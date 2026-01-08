using System;
using System.Collections.Concurrent;
using System.Collections.Generic; 
using System.Text.Json;
using System.Threading.Tasks;
using Lib.Net.Http.WebPush;
using Lib.Net.Http.WebPush.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Push")]
    public class PushController : Controller
    {
        private readonly ILogger<PushController> _log;
        private readonly IConfiguration _cfg;
        private readonly PushServiceClient _client;

        private static readonly ConcurrentDictionary<string, PushSubscriptionDto> _subs
            = new ConcurrentDictionary<string, PushSubscriptionDto>(StringComparer.OrdinalIgnoreCase);

        public PushController(ILogger<PushController> log, IConfiguration cfg)
        {
            _log = log;
            _cfg = cfg;
            _client = new PushServiceClient();
        }

        [HttpGet("VapidPublicKey")]
        [Produces("application/json")]
        public IActionResult VapidPublicKey()
        {
            var pub = (_cfg["WebPush:VapidPublicKey"] ?? "").Trim();
            return Json(new { publicKey = pub });
        }

        [HttpPost("Subscribe")]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public IActionResult Subscribe([FromBody] PushSubscriptionDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.Endpoint)
                || dto.Keys == null
                || string.IsNullOrWhiteSpace(dto.Keys.P256dh)
                || string.IsNullOrWhiteSpace(dto.Keys.Auth))
            {
                return BadRequest(new { ok = false, message = "invalid subscription" });
            }

            var userKey = (User?.Identity?.Name ?? "").Trim();
            if (string.IsNullOrWhiteSpace(userKey))
                return Unauthorized(new { ok = false });

            _subs[userKey] = dto;
            return Ok(new { ok = true });
        }

        [HttpPost("Unsubscribe")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public IActionResult Unsubscribe()
        {
            var userKey = (User?.Identity?.Name ?? "").Trim();
            if (string.IsNullOrWhiteSpace(userKey))
                return Unauthorized(new { ok = false });

            _subs.TryRemove(userKey, out _);
            return Ok(new { ok = true });
        }

        [HttpPost("TestMe")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> TestMe()
        {
            var userKey = (User?.Identity?.Name ?? "").Trim();
            if (string.IsNullOrWhiteSpace(userKey))
                return Unauthorized(new { ok = false });

            if (!_subs.TryGetValue(userKey, out var sub))
                return Ok(new { ok = false, reason = "no-subscription" });

            var payload = new
            {
                title = "E-BOARD",
                body = "테스트 푸시입니다. 클릭하면 E-BOARD로 이동합니다.",
                url = Url.Content("~/Doc/Board"),
                tag = "eboard-test"
            };

            var ok = await SendAsync(sub, payload);
            return Ok(new { ok });
        }

        public async Task<bool> SendToUserAsync(string userKey, string title, string body, string url)
        {
            if (string.IsNullOrWhiteSpace(userKey)) return false;
            if (!_subs.TryGetValue(userKey.Trim(), out var sub)) return false;

            var payload = new { title, body, url, tag = "eboard-approval" };
            return await SendAsync(sub, payload);
        }

        private async Task<bool> SendAsync(PushSubscriptionDto subDto, object payloadObj)
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

                // ✅ PushSubscriptionKeys 타입 사용 금지 → Keys를 Dictionary로 구성
                var subscription = new PushSubscription
                {
                    Endpoint = subDto.Endpoint,
                    Keys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["p256dh"] = subDto.Keys.P256dh,
                        ["auth"] = subDto.Keys.Auth
                    }
                };

                var json = JsonSerializer.Serialize(payloadObj);
                var notification = new PushMessage(json);

                await _client.RequestPushMessageDeliveryAsync(subscription, notification, vapid);
                return true;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Push send failed");
                return false;
            }
        }
    }

    public sealed class PushSubscriptionDto
    {
        public string Endpoint { get; set; } = "";
        public PushKeysDto Keys { get; set; } = new PushKeysDto();
    }

    public sealed class PushKeysDto
    {
        public string P256dh { get; set; } = "";
        public string Auth { get; set; } = "";
    }
}