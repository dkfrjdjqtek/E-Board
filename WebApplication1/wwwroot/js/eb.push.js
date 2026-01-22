
// 2026.01.14 Changed 구독 키(p256dh/auth)를 PushSubscription.getKey()로 추출해 Base64URL로 변환하여 서버로 전송하고 Unregister는 서버 비활성화를 먼저 수행하도록 고정

(function () {
    'use strict';

    const LS_KEY_LAST_ENDPOINT = 'EBPush.LastEndpoint';

    function getCsrfToken() {
        const el = document.querySelector('input[name="__RequestVerificationToken"]');
        return el ? el.value : null;
    }

    function urlBase64ToUint8Array(base64String) {
        const padding = '='.repeat((4 - base64String.length % 4) % 4);
        const base64 = (base64String + padding).replace(/-/g, '+').replace(/_/g, '/');
        const raw = atob(base64);
        const output = new Uint8Array(raw.length);
        for (let i = 0; i < raw.length; ++i) output[i] = raw.charCodeAt(i);
        return output;
    }

    // ArrayBuffer -> Base64URL (Push keys는 이 형태로 보내는 게 안전)
    function arrayBufferToBase64Url(buf) {
        if (!buf) return '';
        const bytes = new Uint8Array(buf);
        let binary = '';
        for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
        const b64 = btoa(binary);
        return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
    }

    async function fetchVapidPublicKey() {
        const r = await fetch('/Push/VapidPublicKey', { credentials: 'same-origin' });
        if (!r.ok) throw new Error('VAPID public key fetch failed: ' + r.status);
        const j = await r.json();
        if (!j || !j.publicKey) throw new Error('VAPID public key missing');
        return j.publicKey;
    }

    async function registerSw() {
        if (!('serviceWorker' in navigator)) throw new Error('ServiceWorker not supported');
        return await navigator.serviceWorker.register('/push/sw.js', { scope: '/push/' });
    }

    async function ensurePermission() {
        if (!('Notification' in window)) throw new Error('Notification not supported');
        if (Notification.permission === 'granted') return true;
        const p = await Notification.requestPermission();
        return p === 'granted';
    }

    async function getSubscriptionSafe(reg) {
        try {
            return await reg.pushManager.getSubscription();
        } catch {
            return null;
        }
    }

    function rememberEndpoint(endpoint) {
        try {
            if (endpoint) localStorage.setItem(LS_KEY_LAST_ENDPOINT, endpoint);
        } catch { }
    }

    function loadRememberedEndpoint() {
        try {
            return (localStorage.getItem(LS_KEY_LAST_ENDPOINT) || '').trim();
        } catch {
            return '';
        }
    }

    async function postJson(url, bodyObj, opts) {
        const headers = Object.assign({ 'Content-Type': 'application/json' }, (opts && opts.headers) ? opts.headers : {});

        const token = getCsrfToken();
        if (token && !headers['RequestVerificationToken']) headers['RequestVerificationToken'] = token;

        const r = await fetch(url, {
            method: 'POST',
            credentials: 'same-origin',
            headers,
            body: JSON.stringify(bodyObj || {})
        });

        let text = '';
        try { text = await r.text(); } catch { }

        let json = null;
        try { json = text ? JSON.parse(text) : null; } catch { }

        return { ok: r.ok, status: r.status, text, json };
    }

    function extractKeys(sub) {
        // ✅ 표준 방식: getKey('p256dh'), getKey('auth')
        // 반환값: ArrayBuffer
        try {
            const p256dhBuf = sub && sub.getKey ? sub.getKey('p256dh') : null;
            const authBuf = sub && sub.getKey ? sub.getKey('auth') : null;

            const p256dh = arrayBufferToBase64Url(p256dhBuf);
            const auth = arrayBufferToBase64Url(authBuf);

            return { p256dh, auth };
        } catch {
            return { p256dh: '', auth: '' };
        }
    }

    async function ensureSubscribed(options) {
        const sendToServer = !!(options && options.sendToServer);

        const ok = await ensurePermission();
        if (!ok) throw new Error('Notification permission denied');

        const reg = await registerSw();

        let sub = await getSubscriptionSafe(reg);
        if (!sub) {
            const publicKey = await fetchVapidPublicKey();
            sub = await reg.pushManager.subscribe({
                userVisibleOnly: true,
                applicationServerKey: urlBase64ToUint8Array(publicKey)
            });
        }

        const endpoint = (sub && sub.endpoint) ? String(sub.endpoint) : '';
        rememberEndpoint(endpoint);

        // 키 추출
        const keys = extractKeys(sub);

        if (sendToServer) {
            const payload = {
                Endpoint: endpoint,
                Keys: {
                    P256dh: keys.p256dh,
                    Auth: keys.auth
                },
                UserAgent: navigator.userAgent
            };

            // 디버그: 서버로 실제 보내는 payload 확인
            // console.log('[EBPush] Subscribe payload=', payload);

            const res = await postJson('/Push/Subscribe', payload);
            if (!res.ok) {
                console.error('[EBPush] Subscribe server failed:', res.status, res.json || res.text);
                return { ok: true, subscribed: true, serverSaved: false, scope: reg.scope, endpoint };
            }

            return { ok: true, subscribed: true, serverSaved: true, scope: reg.scope, endpoint };
        }

        return { ok: true, subscribed: true, serverSaved: false, scope: reg.scope, endpoint };
    }

    async function ping() {
        const reg = await registerSw();
        const sub = await getSubscriptionSafe(reg);
        const endpoint = (sub && sub.endpoint) ? String(sub.endpoint) : loadRememberedEndpoint();

        if (!endpoint) throw new Error('Ping failed: no endpoint');

        const res = await postJson('/Push/Ping', { Endpoint: endpoint });
        if (!res.ok) {
            console.error('[EBPush] Ping failed:', res.status, res.json || res.text);
            throw new Error('Ping failed: ' + res.status);
        }

        return (res.json || { ok: true });
    }

    async function unregister(options) {
        const unsubscribeClient = !!(options && options.unsubscribeClient);

        const reg = await registerSw();
        const sub = await getSubscriptionSafe(reg);

        // ✅ 서버 정리를 먼저 한다(중요)
        const endpoint = (sub && sub.endpoint) ? String(sub.endpoint) : loadRememberedEndpoint();

        if (!endpoint) {
            console.warn('[EBPush] Unregister: no endpoint (already unsubscribed and no saved endpoint)');
            if (unsubscribeClient && sub) {
                try { await sub.unsubscribe(); } catch { }
            }
            return { ok: true, serverSaved: false, unsubscribed: !!unsubscribeClient };
        }

        const res = await postJson('/Push/Unregister', { Endpoint: endpoint });
        if (!res.ok) {
            console.error('[EBPush] Unregister server failed:', res.status, res.json || res.text);
        } else {
            console.log('[EBPush] Unregister server ok:', res.json || res.text);
        }

        let unsubscribed = false;
        if (unsubscribeClient && sub) {
            try {
                unsubscribed = await sub.unsubscribe();
            } catch {
                unsubscribed = false;
            }
        }

        return { ok: true, serverSaved: res.ok, status: res.status, unsubscribed, endpoint };
    }

    async function debugListRegs() {
        if (!('serviceWorker' in navigator)) return [];
        const regs = await navigator.serviceWorker.getRegistrations();
        return regs.map(r => ({ scope: r.scope, active: !!r.active, scriptURL: r.active ? r.active.scriptURL : null }));
    }

    window.EBPush = {
        ensureSubscribed,
        unregister,
        ping,
        debugListRegs
    };

    console.log('EBPush ready. Try: await EBPush.ensureSubscribed({sendToServer:true})');
})();
