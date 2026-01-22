// 2026.01.19 Changed: UserId별 TOP 1 발송을 전체 활성 구독 다건 발송으로 변경하고 WebPushException 타입 의존을 제거하여 컴파일 오류를 해소함
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using Lib.Net.Http.WebPush;
using Lib.Net.Http.WebPush.Authentication;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace WebApplication1.Services
{
    public sealed class WebPushNotifier : IWebPushNotifier
    {
        private readonly ILogger<WebPushNotifier> _log;
        private readonly IConfiguration _cfg;
        private readonly PushServiceClient _client;

        public class WebPushSubscription
        {
            public long Id { get; set; }

            public string? UserId { get; set; }          // AspNetUsers.Id (FK 생성 금지)
            public string? Email { get; set; }

            public string Endpoint { get; set; } = "";   // NOT NULL
            public string P256dh { get; set; } = "";     // NOT NULL
            public string Auth { get; set; } = "";       // NOT NULL

            public string? UserAgent { get; set; }

            public bool IsActive { get; set; } = true;

            public DateTime CreatedAt { get; set; }
            public DateTime UpdatedAt { get; set; }
            public DateTime? LastSeenAt { get; set; }
        }

        public WebPushNotifier(ILogger<WebPushNotifier> log, IConfiguration cfg)
        {
            _log = log;
            _cfg = cfg;
            _client = new PushServiceClient();
        }

        public async Task<bool> SendToUserIdAsync(string userId, string title, string body, string url, string? tag = null)
        {
            userId = (userId ?? "").Trim();
            if (string.IsNullOrWhiteSpace(userId)) return false;

            var subs = await LoadActiveSubscriptionsAsync(userId);
            if (subs == null || subs.Count == 0) return false;

            var payload = new
            {
                title = title ?? "E-BOARD",
                body = body ?? "",
                url = url ?? "/",
                tag = (tag ?? "").Trim()
            };

            bool anyOk = false;

            foreach (var sub in subs)
            {
                var ok = await SendOneAsync(sub, payload);
                if (ok) anyOk = true;
            }

            return anyOk;
        }

        private sealed class DbSub
        {
            public long Id { get; set; }
            public string Endpoint { get; set; } = "";
            public string P256dh { get; set; } = "";
            public string Auth { get; set; } = "";
        }

        private async Task<List<DbSub>> LoadActiveSubscriptionsAsync(string userId)
        {
            var list = new List<DbSub>();

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT Id, Endpoint, P256dh, Auth
FROM dbo.WebPushSubscriptions
WHERE UserId = @UserId
  AND IsActive = 1
ORDER BY UpdatedAt DESC, Id DESC;";
                cmd.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.NVarChar, 64) { Value = userId });

                await using var r = await cmd.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    list.Add(new DbSub
                    {
                        Id = r.GetInt64(0),
                        Endpoint = r.GetString(1),
                        P256dh = r.GetString(2),
                        Auth = r.GetString(3)
                    });
                }

                return list;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "WebPush LoadActiveSubscriptions failed userId={userId}", userId);
                return list;
            }
        }

        private async Task MarkInactiveAsync(long subId)
        {
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
UPDATE dbo.WebPushSubscriptions
   SET IsActive = 0,
       UpdatedAt = SYSUTCDATETIME()
 WHERE Id = @Id;";
                cmd.Parameters.Add(new SqlParameter("@Id", System.Data.SqlDbType.BigInt) { Value = subId });
                await cmd.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "WebPush MarkInactive failed subId={subId}", subId);
            }
        }

        private async Task<bool> SendOneAsync(DbSub sub, object payloadObj)
        {
            try
            {
                var subject = (_cfg["WebPush:VapidSubject"] ?? "").Trim();
                var pub = (_cfg["WebPush:VapidPublicKey"] ?? "").Trim();
                var pri = (_cfg["WebPush:VapidPrivateKey"] ?? "").Trim();

                if (string.IsNullOrWhiteSpace(subject) ||
                    string.IsNullOrWhiteSpace(pub) ||
                    string.IsNullOrWhiteSpace(pri))
                {
                    _log.LogError("WebPush VAPID 설정 누락");
                    return false;
                }

                var vapid = new VapidAuthentication(pub, pri) { Subject = subject };

                var subscription = new PushSubscription
                {
                    Endpoint = sub.Endpoint,
                    Keys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["p256dh"] = sub.P256dh,
                        ["auth"] = sub.Auth
                    }
                };

                var json = JsonSerializer.Serialize(payloadObj);
                var notification = new PushMessage(json);

                await _client.RequestPushMessageDeliveryAsync(subscription, notification, vapid);
                return true;
            }
            catch (Exception ex)
            {
                // 라이브러리 버전에 따라 WebPushException 타입이 없을 수 있어
                // 타입명을 문자열로 검사하여 404/410 계열이면 비활성화 처리
                try
                {
                    var t = ex.GetType();
                    var name = t.Name ?? "";

                    int? statusCode = null;

                    // PushServiceClientException 등에서 StatusCode 읽기 시도
                    var pStatus = t.GetProperty("StatusCode");
                    if (pStatus != null)
                    {
                        var v = pStatus.GetValue(ex);
                        if (v is int i) statusCode = i;
                        else if (v != null && int.TryParse(v.ToString(), out var j)) statusCode = j;
                    }

                    // ResponseStatusCode 등 다른 이름도 대비
                    if (statusCode == null)
                    {
                        var p2 = t.GetProperty("ResponseStatusCode");
                        if (p2 != null)
                        {
                            var v = p2.GetValue(ex);
                            if (v is int i2) statusCode = i2;
                            else if (v != null && int.TryParse(v.ToString(), out var j2)) statusCode = j2;
                        }
                    }

                    // 메시지에 404/410이 포함되는 경우도 대비
                    var msg = (ex.Message ?? "");
                    bool looksGone =
                        (statusCode == 404 || statusCode == 410) ||
                        msg.Contains("410") || msg.Contains("404");

                    _log.LogWarning(ex, "Push send failed subId={subId} type={type} status={status}",
                        sub.Id, name, statusCode.HasValue ? statusCode.Value.ToString() : "");

                    if (looksGone)
                        await MarkInactiveAsync(sub.Id);
                }
                catch
                {
                    // 여기서 예외가 나도 발송 실패 처리만 하면 됨
                    _log.LogError(ex, "Push send failed (secondary handling failed) subId={subId}", sub.Id);
                }

                return false;
            }
        }
    }
}
