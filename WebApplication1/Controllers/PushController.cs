// 2026.01.14 Changed: 브라우저 PushSubscription 원본 JSON(endpoint keys p256dh auth) 그대로 수신하도록 DTO를 통일하고 Subscribe Ping Unregister 바인딩 불일치를 제거

using System.Security.Claims;
using System.Text.Json;
using Lib.Net.Http.WebPush;
using Lib.Net.Http.WebPush.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Data;

namespace WebApplication1.Controllers
{
    [Authorize]
    [IgnoreAntiforgeryToken] // Push API는 JSON fetch 기반이며 Authorize로 보호
    [Route("Push")]
    public class PushController : Controller
    {
        private readonly IConfiguration _cfg;
        private readonly ApplicationDbContext _db;

        public PushController(IConfiguration cfg, ApplicationDbContext db)
        {
            _cfg = cfg;
            _db = db;
        }

        // ------------------------------------------------------------
        // 브라우저가 주는 PushSubscription JSON 형태 그대로 받는다
        // {
        //   endpoint: "...",
        //   expirationTime: null,
        //   keys: { p256dh: "...", auth: "..." }
        // }
        // ------------------------------------------------------------
        public sealed class BrowserPushSubscriptionDto
        {
            public string? endpoint { get; set; }
            public BrowserPushKeysDto? keys { get; set; }
            public object? expirationTime { get; set; }

            // 선택: 클라에서 userAgent를 같이 보내도 되고, 안 보내면 서버에서 Request 헤더로 보강
            public string? userAgent { get; set; }
        }

        public sealed class BrowserPushKeysDto
        {
            public string? p256dh { get; set; }
            public string? auth { get; set; }
        }

        public sealed class SendToUserDto
        {
            public string? TargetUserId { get; set; }
            public string? TargetEmail { get; set; }
            public string? Url { get; set; }
            public string? Tag { get; set; }
            public string? Title { get; set; } // 선택
            public string? Body { get; set; }  // 선택
        }

        // ------------------------------------------------------------
        // GET /Push/VapidPublicKey
        // ------------------------------------------------------------
        [HttpGet("VapidPublicKey")]
        [Produces("application/json")]
        public IActionResult VapidPublicKey()
        {
            var pub = (_cfg["WebPush:VapidPublicKey"] ?? "").Trim();
            if (string.IsNullOrWhiteSpace(pub) || pub.StartsWith("PUT_", StringComparison.OrdinalIgnoreCase))
                return BadRequest(new { ok = false, message = "VAPID PublicKey가 설정되지 않았습니다(WebPush:VapidPublicKey)." });

            return Json(new { ok = true, publicKey = pub });
        }

        // ------------------------------------------------------------
        // POST /Push/Subscribe
        // Endpoint 기준 Upsert + 소유권(UserId/Email) 갱신
        // ------------------------------------------------------------
        [HttpPost("Subscribe")]
        [Produces("application/json")]
        public async Task<IActionResult> Subscribe([FromBody] BrowserPushSubscriptionDto dto)
        {
            var userId = GetUserId();
            var email = GetUserEmail();

            if (string.IsNullOrWhiteSpace(userId) && string.IsNullOrWhiteSpace(email))
                return Unauthorized(new { ok = false, message = "로그인 사용자 정보를 확인할 수 없습니다." });

            var endpoint = (dto?.endpoint ?? "").Trim();
            var p256dh = (dto?.keys?.p256dh ?? "").Trim();
            var auth = (dto?.keys?.auth ?? "").Trim();

            if (string.IsNullOrWhiteSpace(endpoint) || string.IsNullOrWhiteSpace(p256dh) || string.IsNullOrWhiteSpace(auth))
                return BadRequest(new { ok = false, message = "구독 정보(endpoint/keys)가 비었습니다." });

            var ua = (dto?.userAgent ?? "").Trim();
            if (string.IsNullOrWhiteSpace(ua))
                ua = (Request.Headers["User-Agent"].ToString() ?? "").Trim();

            var now = DateTime.UtcNow;

            var row = await _db.WebPushSubscriptions
                .FirstOrDefaultAsync(x => x.Endpoint == endpoint);

            if (row == null)
            {
                row = new WebPushSubscription
                {
                    UserId = userId,
                    Email = email,
                    Endpoint = endpoint,
                    P256dh = p256dh,
                    Auth = auth,
                    UserAgent = string.IsNullOrWhiteSpace(ua) ? null : Trunc(ua, 512),
                    IsActive = true,
                    CreatedAt = now,
                    UpdatedAt = now,
                    LastSeenAt = now
                };
                _db.WebPushSubscriptions.Add(row);
            }
            else
            {
                row.UserId = userId;
                row.Email = email;
                row.P256dh = p256dh;
                row.Auth = auth;
                row.UserAgent = string.IsNullOrWhiteSpace(ua) ? row.UserAgent : Trunc(ua, 512);
                row.IsActive = true;
                row.UpdatedAt = now;
                row.LastSeenAt = now;
            }

            await _db.SaveChangesAsync();

            return Ok(new { ok = true, message = "구독 저장 완료", lastSeenUtc = now });
        }

        // ------------------------------------------------------------
        // POST /Push/Ping
        // 현재 endpoint의 LastSeenAt 갱신 + 소유권 보정
        // ------------------------------------------------------------
        [HttpPost("Ping")]
        [Produces("application/json")]
        public async Task<IActionResult> Ping([FromBody] BrowserPushSubscriptionDto dto)
        {
            var userId = GetUserId();
            var email = GetUserEmail();

            if (string.IsNullOrWhiteSpace(userId) && string.IsNullOrWhiteSpace(email))
                return Unauthorized(new { ok = false, message = "로그인 사용자 정보를 확인할 수 없습니다." });

            var endpoint = (dto?.endpoint ?? "").Trim();
            if (string.IsNullOrWhiteSpace(endpoint))
                return BadRequest(new { ok = false, message = "endpoint required" });

            var now = DateTime.UtcNow;

            var row = await _db.WebPushSubscriptions
                .FirstOrDefaultAsync(x => x.Endpoint == endpoint);

            if (row == null)
                return Ok(new { ok = true, touched = false, reason = "not_found" });

            // 공용PC 정책: 핑에서도 소유권을 현재 로그인 사용자로 보정
            row.UserId = userId;
            row.Email = email;
            row.IsActive = true;
            row.UpdatedAt = now;
            row.LastSeenAt = now;

            await _db.SaveChangesAsync();

            return Ok(new { ok = true, touched = true, lastSeenUtc = now });
        }

        // ------------------------------------------------------------
        // POST /Push/Unregister
        // 로그아웃 시 현재 endpoint만 IsActive=0 처리 (멱등)
        // ------------------------------------------------------------
        [HttpPost("Unregister")]
        [Produces("application/json")]
        public async Task<IActionResult> Unregister([FromBody] BrowserPushSubscriptionDto dto)
        {
            var endpoint = (dto?.endpoint ?? "").Trim();
            if (string.IsNullOrWhiteSpace(endpoint))
                return BadRequest(new { ok = false, message = "endpoint required" });

            var now = DateTime.UtcNow;

            var row = await _db.WebPushSubscriptions
                .Where(x => x.Endpoint == endpoint)
                .OrderByDescending(x => x.Id)
                .FirstOrDefaultAsync();

            if (row == null)
                return Ok(new { ok = true, deactivated = false, reason = "not_found" });

            row.IsActive = false;
            row.UpdatedAt = now;
            row.LastSeenAt = now;

            await _db.SaveChangesAsync();

            return Ok(new { ok = true, deactivated = true });
        }

        // ------------------------------------------------------------
        // POST /Push/SendToUser
        // ------------------------------------------------------------
        [HttpPost("SendToUser")]
        [Produces("application/json")]
        public async Task<IActionResult> SendToUser([FromBody] SendToUserDto dto)
        {
            var targetUserId = (dto.TargetUserId ?? "").Trim();
            var targetEmail = (dto.TargetEmail ?? "").Trim();

            if (string.IsNullOrWhiteSpace(targetUserId) && string.IsNullOrWhiteSpace(targetEmail))
                return BadRequest(new { ok = false, message = "TargetUserId 또는 TargetEmail이 필요합니다." });

            var subject = (_cfg["WebPush:VapidSubject"] ?? "").Trim();
            var pub = (_cfg["WebPush:VapidPublicKey"] ?? "").Trim();
            var pri = (_cfg["WebPush:VapidPrivateKey"] ?? "").Trim();

            if (string.IsNullOrWhiteSpace(subject) || string.IsNullOrWhiteSpace(pub) || string.IsNullOrWhiteSpace(pri))
                return BadRequest(new { ok = false, message = "VAPID 설정(WebPush)이 비었습니다." });

            var now = DateTime.UtcNow;
            var aliveCutoff = now.AddDays(-30);

            var query = _db.WebPushSubscriptions.AsNoTracking().Where(x => x.IsActive);

            if (!string.IsNullOrWhiteSpace(targetUserId))
                query = query.Where(x => x.UserId == targetUserId);

            if (!string.IsNullOrWhiteSpace(targetEmail))
                query = query.Where(x => x.Email == targetEmail);

            query = query.Where(x => x.LastSeenAt == null || x.LastSeenAt >= aliveCutoff);

            var subs = await query
                .OrderByDescending(x => x.LastSeenAt ?? x.UpdatedAt)
                .ToListAsync();

            if (subs.Count == 0)
                return Ok(new { ok = true, sent = 0, message = "활성 구독이 없습니다." });

            var payload = new
            {
                title = string.IsNullOrWhiteSpace(dto.Title) ? "E-BOARD" : dto.Title,
                body = string.IsNullOrWhiteSpace(dto.Body) ? "새 알림이 있습니다." : dto.Body,
                url = string.IsNullOrWhiteSpace(dto.Url) ? "/" : dto.Url,
                tag = string.IsNullOrWhiteSpace(dto.Tag) ? ("eboard-" + Guid.NewGuid().ToString("N")) : dto.Tag
            };

            var message = new PushMessage(JsonSerializer.Serialize(payload)) { TimeToLive = 60 * 60 };

            var authn = new VapidAuthentication(pub, pri) { Subject = subject };
            var client = new PushServiceClient();

            var sent = 0;
            var gone = 0;
            var deactivateEndpoints = new List<string>();

            foreach (var s in subs)
            {
                var pushSub = new PushSubscription
                {
                    Endpoint = s.Endpoint,
                    Keys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["p256dh"] = s.P256dh,
                        ["auth"] = s.Auth
                    }
                };

                try
                {
                    await client.RequestPushMessageDeliveryAsync(pushSub, message, authn);
                    sent++;
                }
                catch (PushServiceClientException ex)
                {
                    var code = (int)ex.StatusCode;
                    if (code == 410 || code == 404)
                    {
                        gone++;
                        deactivateEndpoints.Add(s.Endpoint);
                    }
                }
            }

            if (deactivateEndpoints.Count > 0)
            {
                var rows = await _db.WebPushSubscriptions
                    .Where(x => deactivateEndpoints.Contains(x.Endpoint))
                    .ToListAsync();

                foreach (var r in rows)
                {
                    r.IsActive = false;
                    r.UpdatedAt = DateTime.UtcNow;
                }

                await _db.SaveChangesAsync();
            }

            return Ok(new { ok = true, sent, gone, total = subs.Count });
        }

        private string? GetUserId()
        {
            return User.FindFirstValue(ClaimTypes.NameIdentifier) ?? User.FindFirstValue("sub");
        }

        private string? GetUserEmail()
        {
            var email =
                User.FindFirstValue(ClaimTypes.Email)
                ?? User.FindFirstValue("preferred_username")
                ?? User.Identity?.Name;

            return string.IsNullOrWhiteSpace(email) ? null : email.Trim();
        }

        private static string Trunc(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Length <= max ? s : s.Substring(0, max);
        }
    }

    // ------------------------------------------------------------
    // 주의: 프로젝트에 이미 엔티티가 있으면 이 클래스는 중복이므로 제거
    // ------------------------------------------------------------
    public sealed class WebPushSubscription
    {
        public long Id { get; set; }
        public string? UserId { get; set; }
        public string? Email { get; set; }
        public string Endpoint { get; set; } = "";
        public string P256dh { get; set; } = "";
        public string Auth { get; set; } = "";
        public string? UserAgent { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
        public DateTime? LastSeenAt { get; set; }
    }

    /*
    File: Data/ApplicationDbContext.cs
    // 2026.01.14 Changed: WebPushSubscriptions DbSet 추가 (기존에 이미 있으면 중복 추가 금지)
    public DbSet<WebPushSubscription> WebPushSubscriptions { get; set; } = default!;
    */
}
