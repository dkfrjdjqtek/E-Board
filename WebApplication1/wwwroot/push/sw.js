// 2026.01.23 Changed: title/body가 리소스 키로 오면 서버(/Push/ResolveResourceText)에서 값으로 해석해 표시하도록 추가하고, 실패 시 기존 동작(키 그대로 표시)으로 폴백
// 2026.01.12 Changed: 푸시 알림이 사용자가 닫을 때까지 유지되도록 requireInteraction 적용하고 고정 태그 알림 재알림을 위해 renotify를 적용

self.addEventListener('push', function (event) {
    let data = {};
    try {
        // payload가 JSON이 아닌 경우도 있으니 방어
        if (event.data) {
            try {
                data = event.data.json();
            } catch {
                data = { body: event.data.text() };
            }
        }
    } catch {
        data = {};
    }

    const titleRaw = data.title || 'E-BOARD';
    const bodyRaw = data.body || '';
    const url = data.url || '/';

    // 디버깅: tag가 고정이면 알림이 교체되어 "한 번만 뜨는 것처럼" 보일 수 있음
    // 운영에서는 tag를 쓰고 싶으면 data.tag를 주고, 디버깅 중엔 tag를 비워도 됩니다.
    const tag = (typeof data.tag === 'string' && data.tag.trim()) ? data.tag.trim() : '';

    const options = {
        body: bodyRaw,
        data: { url: url },
        requireInteraction: true // 운영: 사용자가 닫거나 클릭할 때까지 알림 유지
        // 필요 시 icon, badge 추가 가능
    };

    // A안: 문서 1건당 알림 1개(고정 tag)로 유지
    // 30분마다 같은 tag로 다시 보내면 누적이 아니라 기존 알림이 교체되지만,
    // renotify를 켜면 교체 시에도 다시 알림이 울리도록 시도합니다(환경에 따라 다를 수 있음).
    if (tag) {
        options.tag = tag;
        options.renotify = true;
        options.requireInteraction = true //사용자가 직접 닫거나 클릭할 때까지 유지
    }

    // ------------------------------------------------------------
    // (추가) 다국어: title/body가 "리소스 키"로 오면 서버에서 값으로 해석
    // - 서버가 키를 보내는 구조(예: PUSH_Approval_Next_Title)를 유지하면서
    //   사용자 브라우저 언어( navigator.language ) 기준으로 값을 받아 표시한다.
    // - 실패하면 기존대로 키를 그대로 보여준다(진단용).
    // ------------------------------------------------------------
    function looksLikeResourceKey(s) {
        if (!s) return false;
        if (typeof s !== 'string') return false;
        const x = s.trim();
        if (!x) return false;
        // PUSH_, DOC_, COMMON_ 등 리소스 키 패턴에 맞추어 최소한만 판정
        if (x.startsWith('PUSH_')) return true;
        if (x.startsWith('DOC_')) return true;
        if (x.startsWith('COMMON_')) return true;
        // SummaryTitle 같은 케이스도 키로 취급(원하면 여기에 추가)
        if (x === 'SummaryTitle') return true;
        return false;
    }

    async function resolveResourceTextIfNeeded() {
        const titleIsKey = looksLikeResourceKey(titleRaw);
        const bodyIsKey = looksLikeResourceKey(bodyRaw);

        if (!titleIsKey && !bodyIsKey) {
            return { title: titleRaw, body: bodyRaw };
        }

        // args가 payload에 있으면 같이 전달 (서버에서 string.Format 처리)
        const args = (data && typeof data.args === 'object' && data.args) ? data.args : null;

        // 브라우저 언어를 서버로 전달 (예: ko-KR)
        let lang = '';
        try {
            lang = (self.navigator && self.navigator.language) ? String(self.navigator.language) : '';
        } catch { lang = ''; }

        try {
            const r = await fetch('/Push/ResolveResourceText', {
                method: 'POST',
                credentials: 'include',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    titleKey: titleIsKey ? titleRaw : null,
                    bodyKey: bodyIsKey ? bodyRaw : null,
                    args: args,
                    lang: lang
                })
            });

            if (!r.ok) {
                return { title: titleRaw, body: bodyRaw };
            }

            const j = await r.json();
            const t = (j && typeof j.title === 'string' && j.title.trim()) ? j.title : titleRaw;
            const b = (j && typeof j.body === 'string') ? j.body : bodyRaw;
            return { title: t, body: b };
        } catch {
            return { title: titleRaw, body: bodyRaw };
        }
    }

    event.waitUntil((async () => {
        const resolved = await resolveResourceTextIfNeeded();

        // options.body는 원본이 들어가 있으니, 해석된 body로 교체
        options.body = resolved.body || '';

        await self.registration.showNotification(resolved.title || 'E-BOARD', options);
    })());
});

self.addEventListener('notificationclick', function (event) {
    event.notification.close();

    const url =
        (event.notification && event.notification.data && event.notification.data.url)
            ? event.notification.data.url
            : '/';

    event.waitUntil((async () => {
        const allClients = await clients.matchAll({ type: 'window', includeUncontrolled: true });

        for (const c of allClients) {
            try {
                const u = new URL(c.url);
                if (u.origin === self.location.origin) {
                    await c.focus();
                    await c.navigate(url);
                    return;
                }
            } catch { }
        }

        await clients.openWindow(url);
    })());
});
