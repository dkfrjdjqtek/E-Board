// 2025.11.10 Changed: 미사용 IStringLocalizer 필드 및 생성자 인수 제거로 IDE0052 경고 해소 기타 로직 변경 없음
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc/Comments2")]
    public class DocCommentsController : Controller
    {
        // 다른 네임스페이스와 충돌 방지를 위해 완전 수식 사용
        private readonly WebApplication1.Services.IAuditLogger _audit;

        // 상수 배열(경고 억제용)
        private static readonly string[] ERR_INVALID_PAYLOAD = { "DOC_Err_InvalidPayload" };
        private static readonly string[] ERR_DOC_NOT_FOUND = { "DOC_Err_DocumentNotFound" };
        private static readonly string[] ERR_INVALID_REQUEST = { "DOC_Err_InvalidRequest" };

        public DocCommentsController(WebApplication1.Services.IAuditLogger audit)
        {
            _audit = audit;
        }

        // GET /Doc/Comments?DocID=DOC_xxx
        [HttpGet]
        public IActionResult List([FromQuery] string docId)
        {
            if (string.IsNullOrWhiteSpace(docId))
            {
                return BadRequest(new
                {
                    messages = ERR_DOC_NOT_FOUND,
                    fieldErrors = new Dictionary<string, string[]> { { "DocId", ERR_DOC_NOT_FOUND } }
                });
            }

            // TODO: DB에서 댓글+첨부 목록 로드
            var list = Array.Empty<object>();
            return Json(new { items = list });
        }

        // POST /Doc/Comments  (JSON 본문, CSRF 헤더 사용)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([FromBody] CreateCommentDto dto)
        {
            if (dto is null)
            {
                return BadRequest(new
                {
                    messages = ERR_INVALID_PAYLOAD,
                    fieldErrors = new Dictionary<string, string[]> { { "Payload", ERR_INVALID_PAYLOAD } }
                });
            }

            if (string.IsNullOrWhiteSpace(dto.DocID))
                ModelState.AddModelError("DocId", "DOC_Err_DocumentNotFound");
            if (string.IsNullOrWhiteSpace(dto.Body))
                ModelState.AddModelError("Body", "DOC_Val_Required");

            if (!ModelState.IsValid)
            {
                var fieldErrors = ModelState
                    .Where(kv => kv.Value?.Errors?.Count > 0)
                    .ToDictionary(
                        kv => kv.Key,
                        kv => kv.Value!.Errors.Select(e => e.ErrorMessage).Distinct().ToArray()
                    );
                var summary = fieldErrors.Values.SelectMany(v => v).Distinct().ToArray();
                return BadRequest(new { messages = summary, fieldErrors });
            }

            var actorId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

            // TODO: 댓글 저장 및 새 CommentID 반환
            long newId = 0;

            try
            {
                await _audit.LogAsync(
                    docId: dto.DocID!,
                    actorId: actorId,
                    actionCode: "CommentCreated",
                    detailJson: null
                );
            }
            catch
            {
                // 감사로그 실패는 무시
            }

            return Json(new { commentId = newId, message = "DOC_Msg_CommentCreated" });
        }

        // DELETE /Doc/Comments/{id}?DocID=DOC_xxx  (CSRF 헤더 사용)
        [HttpDelete("{id:long}")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(long id, [FromQuery] string docId)
        {
            if (id <= 0 || string.IsNullOrWhiteSpace(docId))
            {
                return BadRequest(new
                {
                    messages = ERR_INVALID_REQUEST,
                    fieldErrors = new Dictionary<string, string[]>
                    {
                        { "CommentId", ERR_INVALID_REQUEST },
                        { "DocId",     ERR_INVALID_REQUEST }
                    }
                });
            }

            // TODO: 권한 검증 및 소프트 삭제 수행

            var actorId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            try
            {
                await _audit.LogAsync(
                    docId: docId,
                    actorId: actorId,
                    actionCode: "CommentDeleted",
                    detailJson: null
                );
            }
            catch
            {
                // 감사로그 실패는 무시
            }

            return Json(new { message = "DOC_Msg_CommentDeleted" });
        }

        // POST /Doc/Comments/attach  (JSON 본문, CSRF 헤더 사용)
        [HttpPost("attach")]
        [ValidateAntiForgeryToken]
        public IActionResult Attach([FromBody] AttachDto dto)
        {
            if (dto is null)
            {
                return BadRequest(new
                {
                    messages = ERR_INVALID_PAYLOAD,
                    fieldErrors = new Dictionary<string, string[]> { { "Payload", ERR_INVALID_PAYLOAD } }
                });
            }

            if (string.IsNullOrWhiteSpace(dto.DocID))
                ModelState.AddModelError("DocId", "DOC_Err_DocumentNotFound");
            if (string.IsNullOrWhiteSpace(dto.FileKey))
                ModelState.AddModelError("FileKey", "DOC_Val_Required");
            if (string.IsNullOrWhiteSpace(dto.OriginalName))
                ModelState.AddModelError("OriginalName", "DOC_Val_Required");

            if (!ModelState.IsValid)
            {
                var fieldErrors = ModelState
                    .Where(kv => kv.Value?.Errors?.Count > 0)
                    .ToDictionary(
                        kv => kv.Key,
                        kv => kv.Value!.Errors.Select(e => e.ErrorMessage).Distinct().ToArray()
                    );
                var summary = fieldErrors.Values.SelectMany(v => v).Distinct().ToArray();
                return BadRequest(new { messages = summary, fieldErrors });
            }

            // TODO: 첨부 메타 INSERT
            return Json(new { message = "DOC_Msg_AttachSaved" });
        }

        // ----- DTOs -----
        public class CreateCommentDto
        {
            public string? DocID { get; set; }
            public long? ParentCommentID { get; set; }
            public string? Body { get; set; }
            public List<FileMeta>? Files { get; set; }
        }

        public class AttachDto
        {
            public string? DocID { get; set; }
            public long? CommentID { get; set; }
            public string? FileKey { get; set; }
            public string? OriginalName { get; set; }
            public string? StoragePath { get; set; }
            public string? ContentType { get; set; }
            public long? ByteSize { get; set; }
            public string? SHA256Hex { get; set; }
        }

        public class FileMeta
        {
            public string? FileKey { get; set; }
            public string? OriginalName { get; set; }
            public string? ContentType { get; set; }
            public long? ByteSize { get; set; }
        }
    }
}
