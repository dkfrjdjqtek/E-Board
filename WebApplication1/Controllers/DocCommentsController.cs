// 2025.10.14 Added: 문서 댓글/대댓글/첨부 최소 API 스켈레톤 (EB-VALIDATE 규격 응답)
// 2025.10.14 Added: 목록 조회, 작성(루트/대댓글), 삭제(소프트), 첨부 메타 등록
// 주의 1) 기존 코드/키/구조 변경 없음. 신규 컨트롤러만 추가합니다.
// 주의 2) 실제 DB 저장은 TODO 위치에서 기존 리포지토리/ORM을 사용해 연결하십시오.
// 주의 3) 모든 에러는 Validation Summary/Valid State(EB-VALIDATE)를 따릅니다.
// 주의 4) i18n: 모든 메시지는 리소스 키로 반환합니다(예: DOC_Err_InvalidPayload).

using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Localization;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc/Comments")]
    public class DocCommentsController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IAuditLogger _audit;

        public DocCommentsController(IStringLocalizer<SharedResource> S, IAuditLogger audit)
        {
            _S = S;
            _audit = audit;
        }

        // 2025.10.14 Added: 문서별 댓글 목록 조회 (루트+대댓글 포함, 단순 시간순)
        // GET /Doc/Comments?docId=DOC_xxx
        [HttpGet]
        public IActionResult List([FromQuery] string docId)
        {
            if (string.IsNullOrWhiteSpace(docId))
            {
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_DocumentNotFound" },
                    fieldErrors = new Dictionary<string, string[]> { { "DocId", new[] { "DOC_Err_DocumentNotFound" } } }
                });
            }

            // TODO: DB에서 DocumentComments + (선택) DocumentFiles join하여 목록 로드
            // 반환 형식 예시(키/구조는 자유, 단 i18n 키 사용):
            // [{ commentId, parentCommentId, threadRootId, depth, body, createdBy, createdAt, hasAttachment, files:[{fileId, originalName, byteSize}] }]
            var list = Array.Empty<object>(); // TODO 교체

            return Json(new { items = list });
        }

        // 2025.10.14 Added: 댓글 작성(루트/대댓글)
        // POST /Doc/Comments
        [HttpPost]
        public async Task<IActionResult> Create([FromBody] CreateCommentDto dto)
        {
            if (dto is null)
            {
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_InvalidPayload" },
                    fieldErrors = new Dictionary<string, string[]> { { "Payload", new[] { "DOC_Err_InvalidPayload" } } }
                });
            }

            if (string.IsNullOrWhiteSpace(dto.docId))
                ModelState.AddModelError("DocId", "DOC_Err_DocumentNotFound");
            if (string.IsNullOrWhiteSpace(dto.body))
                ModelState.AddModelError("Body", "DOC_Val_Required");

            if (!ModelState.IsValid)
            {
                var fieldErrors = ModelState
                    .Where(kv => kv.Value?.Errors?.Count > 0)
                    .ToDictionary(kv => kv.Key, kv => kv.Value!.Errors.Select(e => e.ErrorMessage).Distinct().ToArray());
                var summary = fieldErrors.Values.SelectMany(v => v).Distinct().ToArray();
                return BadRequest(new { messages = summary, fieldErrors });
            }

            var actorId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";

            // TODO: DB INSERT
            // - DocumentComments: DocId, ParentCommentId(dto.parentCommentId), ThreadRootId(루트 계산), Depth(0/1+), Body(dto.body), CreatedBy(actorId)
            // - (선택) DocumentFiles: dto.files가 있으면 메타 INSERT (실제 파일 저장/경로는 기존 스토리지 서비스 사용)
            // - 생성된 commentId 반환
            long newId = 0; // TODO 생성된 CommentId로 교체

            // 감사 로그(댓글 작성)
            try { await _audit.LogAsync(dto.docId!, actorId, "CommentCreated", null); } catch { /* ignore */ }

            return Json(new { commentId = newId, message = "DOC_Msg_CommentCreated" });
        }

        // 2025.10.14 Added: 댓글 소프트 삭제
        // DELETE /Doc/Comments/{id}?docId=DOC_xxx
        [HttpDelete("{id:long}")]
        public async Task<IActionResult> Delete(long id, [FromQuery] string docId)
        {
            if (id <= 0 || string.IsNullOrWhiteSpace(docId))
            {
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_InvalidRequest" },
                    fieldErrors = new Dictionary<string, string[]>
                    {
                        { "CommentId", new[] { "DOC_Err_InvalidRequest" } },
                        { "DocId", new[] { "DOC_Err_InvalidRequest" } }
                    }
                });
            }

            // TODO: 소유자/결재권자/공유 권한 검증(권한 없으면 403)
            // TODO: DocumentComments.IsDeleted=1, UpdatedAt=UTCNOW

            var actorId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            try { await _audit.LogAsync(docId, actorId, "CommentDeleted", null); } catch { /* ignore */ }

            return Json(new { message = "DOC_Msg_CommentDeleted" });
        }

        // 2025.10.14 Added: 첨부 메타 등록(파일 업로드는 기존 업로더 사용, 여기는 메타만 연결)
        // POST /Doc/Comments/attach
        [HttpPost("attach")]
        public IActionResult Attach([FromBody] AttachDto dto)
        {
            if (dto is null)
            {
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_InvalidPayload" },
                    fieldErrors = new Dictionary<string, string[]> { { "Payload", new[] { "DOC_Err_InvalidPayload" } } }
                });
            }

            if (string.IsNullOrWhiteSpace(dto.docId))
                ModelState.AddModelError("DocId", "DOC_Err_DocumentNotFound");
            if (string.IsNullOrWhiteSpace(dto.fileKey))
                ModelState.AddModelError("FileKey", "DOC_Val_Required");
            if (string.IsNullOrWhiteSpace(dto.originalName))
                ModelState.AddModelError("OriginalName", "DOC_Val_Required");

            if (!ModelState.IsValid)
            {
                var fieldErrors = ModelState
                    .Where(kv => kv.Value?.Errors?.Count > 0)
                    .ToDictionary(kv => kv.Key, kv => kv.Value!.Errors.Select(e => e.ErrorMessage).Distinct().ToArray());
                var summary = fieldErrors.Values.SelectMany(v => v).Distinct().ToArray();
                return BadRequest(new { messages = summary, fieldErrors });
            }

            // TODO: DocumentFiles INSERT (DocId, CommentId?, FileKey, OriginalName, StoragePath?, ContentType?, ByteSize, Sha256, UploadedBy)
            // - 파일 저장 자체는 기존 업로드 엔드포인트/미들웨어 사용
            // - 여기서는 업로드 완료 후 메타만 기록

            return Json(new { message = "DOC_Msg_AttachSaved" });
        }

        // DTO들: 클라이언트와 키 이름 동일 유지
        public class CreateCommentDto
        {
            public string? docId { get; set; }
            public long? parentCommentId { get; set; }   // 대댓글이면 대상 id
            public string? body { get; set; }
            public List<FileMeta>? files { get; set; }   // (선택) 업로드 완료 파일 메타
        }

        public class AttachDto
        {
            public string? docId { get; set; }
            public long? commentId { get; set; }         // 특정 댓글 첨부 시 지정
            public string? fileKey { get; set; }         // 스토리지 키(유니크)
            public string? originalName { get; set; }
            public string? storagePath { get; set; }
            public string? contentType { get; set; }
            public long? byteSize { get; set; }
            public string? sha256Hex { get; set; }       // 선택
        }

        public class FileMeta
        {
            public string? fileKey { get; set; }
            public string? originalName { get; set; }
            public string? contentType { get; set; }
            public long? byteSize { get; set; }
        }
    }
}
