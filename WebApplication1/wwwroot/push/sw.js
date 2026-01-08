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
        data: { url: url }
        // 필요 시 icon, badge 추가 가능
        // requireInteraction: true // 디버깅용: 자동으로 빨리 사라지는 것처럼 보이면 잠시 켜서 확인
    };

    if (tag) options.tag = tag;

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