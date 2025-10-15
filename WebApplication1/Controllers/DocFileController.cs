// 2025.10.15 Changed: 업로드 Select 반환 형식 불일치로 인한 CS0411 해결(명시 DTO로 통일)
// 2025.10.15 Changed: 파일 저장에 CopyToAsync 사용하여 async 경고(CS1998) 제거
// 2025.10.15 Changed: Download 액션에서 불필요한 async 제거(경고 제거), 나머지 로직/라우팅/검증 불변
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Security.Claims;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Antiforgery;
using Microsoft.Extensions.Configuration;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("[controller]")]
    public sealed class DocFileController : Controller
    {
        private readonly IWebHostEnvironment _env;
        private readonly IConfiguration _cfg;
        private readonly IAntiforgery _antiforgery;

        public DocFileController(IWebHostEnvironment env, IConfiguration cfg, IAntiforgery antiforgery)
        {
            _env = env;
            _cfg = cfg;
            _antiforgery = antiforgery;
        }

        // 2025.10.15 Added: 업로드 결과 통일 DTO (익명형 혼용으로 Select 타입 추론 실패했던 문제 해결)
        private sealed class UploadItemResult
        {
            public string? fileKey { get; init; }
            public string? originalName { get; init; }
            public string? contentType { get; init; }
            public long? byteSize { get; init; }
            public string? error { get; init; }
        }

        // Form: files[] (multiple), optional: docId, commentId
        // Return: { items: [{ fileKey, originalName, contentType, byteSize }], errors: [...] }
        [HttpPost("Upload")]
        [ValidateAntiForgeryToken]
        [RequestSizeLimit(50_000_000)]
        [Consumes("multipart/form-data")]
        [Produces("application/json")]
        public async Task<IActionResult> Upload([FromQuery] string? docId, [FromQuery] long? commentId)
        {
            if (Request.Form?.Files == null || Request.Form.Files.Count == 0)
                return BadRequest(new { messages = new[] { "DOC_File_Err_NoFiles" } });

            var maxSize = _cfg.GetValue<long?>("Files:MaxBytes") ?? 50_000_000L;
            var allowedExt = (_cfg.GetValue<string>("Files:AllowedExtensions") ?? ".pdf,.png,.jpg,.jpeg,.txt,.docx,.xlsx")
                                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                                .Select(x => x.StartsWith('.') ? x.ToLowerInvariant() : "." + x.ToLowerInvariant())
                                .ToArray();

            var storeRoot = Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "uploads");
            Directory.CreateDirectory(storeRoot);

            // 2025.10.15 Changed: 명시 DTO + foreach 로 통일
            var items = new System.Collections.Generic.List<UploadItemResult>();

            foreach (var file in Request.Form.Files)
            {
                var originalName = Path.GetFileName(file.FileName ?? "");
                var ext = Path.GetExtension(originalName).ToLowerInvariant();

                if (!allowedExt.Contains(ext))
                {
                    items.Add(new UploadItemResult { error = "DOC_File_Err_ExtensionNotAllowed", originalName = originalName });
                    continue;
                }
                if (file.Length <= 0)
                {
                    items.Add(new UploadItemResult { error = "DOC_File_Err_Empty", originalName = originalName });
                    continue;
                }
                if (file.Length > maxSize)
                {
                    items.Add(new UploadItemResult { error = "DOC_File_Err_TooLarge", originalName = originalName });
                    continue;
                }

                var fileKey = $"F_{DateTime.UtcNow:yyyyMMddHHmmssfff}_{Guid.NewGuid():N}{ext}";
                var safeFolder = Path.Combine(storeRoot, DateTime.UtcNow.ToString("yyyyMMdd"));
                Directory.CreateDirectory(safeFolder);
                var savePath = Path.Combine(safeFolder, fileKey);

                // 2025.10.15 Changed: 비동기 저장으로 경고 제거
                using (var fs = new FileStream(savePath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                {
                    await file.CopyToAsync(fs);
                }

                items.Add(new UploadItemResult
                {
                    fileKey = fileKey,
                    originalName = originalName,
                    contentType = string.IsNullOrWhiteSpace(file.ContentType) ? "application/octet-stream" : file.ContentType,
                    byteSize = file.Length
                });
            }

            var errs = items.Where(x => x.error != null).ToArray();
            var oks = items.Where(x => x.error == null).ToArray();

            if (oks.Length == 0)
                return BadRequest(new { messages = errs.Select(e => e.error!).Distinct().ToArray() });

            return Json(new { items = oks, errors = errs });
        }

        // 2025.10.15 Changed: async 제거(경고 해소). 스트림은 File()에서 처리.
        [HttpGet("Download/{fileKey}")]
        public IActionResult Download([FromRoute] string fileKey)
        {
            if (string.IsNullOrWhiteSpace(fileKey) || fileKey.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                return BadRequest(new { messages = new[] { "DOC_File_Err_InvalidKey" } });

            var root = Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "uploads");
            if (!Directory.Exists(root))
                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

            string? path = null;
            foreach (var dayDir in Directory.EnumerateDirectories(root))
            {
                var candidate = Path.Combine(dayDir, fileKey);
                if (System.IO.File.Exists(candidate))
                {
                    path = candidate;
                    break;
                }
            }

            if (path == null)
                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

            var contentType = "application/octet-stream";
            var ext = Path.GetExtension(fileKey).ToLowerInvariant();
            if (ext is ".png") contentType = "image/png";
            else if (ext is ".jpg" or ".jpeg") contentType = "image/jpeg";
            else if (ext is ".pdf") contentType = "application/pdf";
            else if (ext is ".txt") contentType = "text/plain";

            var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true);
            return File(stream, contentType, fileKey);
        }

        [HttpDelete("Remove")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public IActionResult Remove([FromQuery] string fileKey)
        {
            if (!User.IsInRole("Admin"))
                return Forbid();

            if (string.IsNullOrWhiteSpace(fileKey))
                return BadRequest(new { messages = new[] { "DOC_File_Err_InvalidKey" } });

            var root = Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "uploads");
            var found = false;
            if (Directory.Exists(root))
            {
                foreach (var dayDir in Directory.EnumerateDirectories(root))
                {
                    var p = Path.Combine(dayDir, fileKey);
                    if (System.IO.File.Exists(p))
                    {
                        System.IO.File.Delete(p);
                        found = true;
                        break;
                    }
                }
            }

            if (!found)
                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

            return Json(new { ok = true });
        }
    }
}
