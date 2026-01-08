
// 2026.01.08 Changed: EBPush ensureSubscribed를 복구하고 debugListRegs debugUnsubscribe를 포함하여 구독과 상태확인이 모두 가능하도록 정리

(function () {
    'use strict';

    const EBPush = (window.EBPush = window.EBPush || {});

    const DEFAULT_SCOPE = '/push/';
    const SW_URL = '/push/sw.js';
    const VAPID_URL = '/Push/VapidPublicKey';

    function _absUrl(u) {
        try { return new URL(u, location.origin).toString(); } catch { return u; }
    }

    async function _fetchJson(url, opts) {
        const r = await fetch(url, Object.assign({ credentials: 'same-origin' }, (opts || {})));
        if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + url);
        return await r.json();
    }

    async function _getVapidPublicKey() {
        const j = await _fetchJson(VAPID_URL);
        if (!j || !j.publicKey || typeof j.publicKey !== 'string') throw new Error('VAPID publicKey 응답이 비어있습니다.');
        return j.publicKey.trim();
    }

    function _base64UrlToUint8Array(base64UrlString) {
        // VAPID 공개키는 base64url 문자열입니다.
        const padding = '='.repeat((4 - (base64UrlString.length % 4)) % 4);
        const base64 = (base64UrlString + padding).replace(/-/g, '+').replace(/_/g, '/');
        const rawData = atob(base64);
        const outputArray = new Uint8Array(rawData.length);
        for (let i = 0; i < rawData.length; ++i) outputArray[i] = rawData.charCodeAt(i);
        return outputArray;
    }

    async function _registerSw() {
        if (!('serviceWorker' in navigator)) throw new Error('ServiceWorker 미지원 브라우저입니다.');

        // 기존 등록이 있으면 재사용
        let reg = await navigator.serviceWorker.getRegistration(DEFAULT_SCOPE);
        if (reg) return reg;

        // 없으면 등록 (중요: scope는 /push/로 고정)
        reg = await navigator.serviceWorker.register(SW_URL, { scope: DEFAULT_SCOPE });
        return reg;
    }

    // ------------------------------------------------------------
    // 핵심: 구독 보장
    // - 권한 요청
    // - /push/ scope로 SW 등록
    // - VAPID 공개키로 subscribe
    // - 서버로 subscription JSON 전송(선택)
    // ------------------------------------------------------------
    EBPush.ensureSubscribed = async function ensureSubscribed(options) {
        options = options || {};
        const sendToServer = (options.sendToServer !== false); // 기본 true
        const subscribeUrl = options.subscribeUrl || '/Push/Subscribe'; // 서버가 받는 엔드포인트가 다르면 여기만 바꾸세요

        if (!('Notification' in window)) return { ok: false, reason: 'no_notification_api' };
        if (!('serviceWorker' in navigator)) return { ok: false, reason: 'no_service_worker' };
        if (!('PushManager' in window)) return { ok: false, reason: 'no_push_manager' };

        // 권한
        if (Notification.permission === 'default') {
            const perm = await Notification.requestPermission();
            if (perm !== 'granted') return { ok: false, reason: 'permission_' + perm };
        } else if (Notification.permission !== 'granted') {
            return { ok: false, reason: 'permission_' + Notification.permission };
        }

        const reg = await _registerSw();

        // 이미 구독이 있으면 그대로 반환
        let sub = await reg.pushManager.getSubscription();
        if (!sub) {
            const vapidPublicKey = await _getVapidPublicKey();
            const appServerKey = _base64UrlToUint8Array(vapidPublicKey);

            sub = await reg.pushManager.subscribe({
                userVisibleOnly: true,
                applicationServerKey: appServerKey
            });
        }

        // 서버로 저장 (서버가 아직 없으면 sendToServer:false로 호출)
        if (sendToServer) {
            try {
                const payload = sub.toJSON ? sub.toJSON() : JSON.parse(JSON.stringify(sub));

                const headers = { 'Content-Type': 'application/json' };
                // 프로젝트에 EB CSRF가 있으면 같이 붙이도록 (없어도 동작)
                const tokenEl = document.querySelector('input[name="__RequestVerificationToken"]');
                if (tokenEl && tokenEl.value) headers['RequestVerificationToken'] = tokenEl.value;

                await fetch(subscribeUrl, {
                    method: 'POST',
                    credentials: 'same-origin',
                    headers,
                    body: JSON.stringify(payload)
                });
            } catch (e) {
                // 서버 저장 실패는 구독 자체 실패가 아니므로 ok 유지 + warn
                console.warn('Push subscription 서버 저장 실패', e);
                return { ok: true, subscribed: true, serverSaved: false, scope: reg.scope, endpoint: sub.endpoint };
            }
        }

        return { ok: true, subscribed: true, serverSaved: true, scope: reg.scope, endpoint: sub.endpoint };
    };

    // ------------------------------------------------------------
    // 디버그: 현재 origin 서비스워커 등록 및 구독 상태 출력
    // ------------------------------------------------------------
    EBPush.debugListRegs = async function debugListRegs() {
        if (!('serviceWorker' in navigator)) {
            console.warn('ServiceWorker 미지원 브라우저입니다.');
            return [];
        }

        const regs = await navigator.serviceWorker.getRegistrations();
        const rows = [];

        for (const r of regs) {
            const active = r.active || r.waiting || r.installing;
            let sub = null;
            let subErr = '';

            try { if (r.pushManager) sub = await r.pushManager.getSubscription(); }
            catch (e) { subErr = (e && e.message) ? e.message : String(e); }

            rows.push({
                scope: r.scope,
                scriptURL: active ? active.scriptURL : '',
                state: active ? active.state : '',
                hasSubscription: !!sub,
                endpointHead: sub && sub.endpoint ? (sub.endpoint.substring(0, 60) + '...') : '',
                subscriptionError: subErr
            });
        }

        console.table(rows);
        return rows;
    };

    // ------------------------------------------------------------
    // 디버그: /push/ 등록의 구독 해제
    // ------------------------------------------------------------
    EBPush.debugUnsubscribe = async function debugUnsubscribe() {
        if (!('serviceWorker' in navigator)) return { ok: false, reason: 'no_serviceWorker' };

        const reg = await navigator.serviceWorker.getRegistration(DEFAULT_SCOPE);
        if (!reg || !reg.pushManager) return { ok: false, reason: 'no_registration' };

        const sub = await reg.pushManager.getSubscription();
        if (!sub) return { ok: true, unsubscribed: false, reason: 'no_subscription' };

        const ok = await sub.unsubscribe();
        return { ok: true, unsubscribed: ok };
    };

})();
