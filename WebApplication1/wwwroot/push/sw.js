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

    const title = data.title || 'E-BOARD';
    const body = data.body || '';
    const url = data.url || '/';

    // 디버깅: tag가 고정이면 알림이 교체되어 "한 번만 뜨는 것처럼" 보일 수 있음
    // 운영에서는 tag를 쓰고 싶으면 data.tag를 주고, 디버깅 중엔 tag를 비워도 됩니다.
    const tag = (typeof data.tag === 'string' && data.tag.trim()) ? data.tag.trim() : '';

    const options = {
        body: body,
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

    event.waitUntil(self.registration.showNotification(title, options));
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
