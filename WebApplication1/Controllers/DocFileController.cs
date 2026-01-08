// 2025.12.26 Changed: 업로드 시 모든 이미지 확장자를 허용하도록 기본 허용 확장자에 이미지 계열을 추가하고 업로드 완료 후 dbo DocumentFiles 에 파일 메타를 즉시 저장하여 상신 후에도 첨부 정보가 누락되지 않도록 보완

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
using Microsoft.Data.SqlClient;
using System.Data;
using System.Security.Cryptography;

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

        private sealed class UploadItemResult
        {
            public string? fileKey { get; init; }
            public string? originalName { get; init; }
            public string? contentType { get; init; }
            public long? byteSize { get; init; }
            public string? error { get; init; }
        }

        [HttpPost("Upload")]
        [ValidateAntiForgeryToken]
        [RequestSizeLimit(50_000_000)]
        [Consumes("multipart/form-data")]
        [Produces("application/json")]
        public async Task<IActionResult> Upload([FromQuery] string? docId, [FromQuery] long? commentId)
        {
            return StatusCode(410, new
            {
                messages = new[] { "DOC_Err_UploadFailed" },
                detail = "DocFile.Upload disabled. Use /Doc/Upload."
            });

            if (Request.Form?.Files == null || Request.Form.Files.Count == 0)
                return BadRequest(new { messages = new[] { "DOC_File_Err_NoFiles" } });
             
            // 정책: docId 없는 Upload 호출은 절대 허용하지 않음(파일 저장/DB 저장 전 차단)
            if (string.IsNullOrWhiteSpace(docId))
                return BadRequest(new { messages = new[] { "DOC_File_Err_DocIdRequired" } });

            var maxSize = _cfg.GetValue<long?>("Files:MaxBytes") ?? 50_000_000L;

            // 기본 허용 확장자: 기존 + 모든 이미지 확장자 계열 추가
            var defaultAllowed = ".pdf,.txt,.docx,.xlsx"
                               + ",.png,.jpg,.jpeg,.gif,.bmp,.webp,.tif,.tiff,.svg,.ico,.heic,.heif";

            var allowedExt = (_cfg.GetValue<string>("Files:AllowedExtensions") ?? defaultAllowed)
                                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                                .Select(x => x.StartsWith('.') ? x.ToLowerInvariant() : "." + x.ToLowerInvariant())
                                .Distinct()
                                .ToArray();

            var storeRoot = Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "uploads");
            Directory.CreateDirectory(storeRoot);

            var items = new System.Collections.Generic.List<UploadItemResult>();

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            var hasDb = !string.IsNullOrWhiteSpace(cs);

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
                var dayFolderName = DateTime.UtcNow.ToString("yyyyMMdd");
                var safeFolder = Path.Combine(storeRoot, dayFolderName);
                Directory.CreateDirectory(safeFolder);

                var savePath = Path.Combine(safeFolder, fileKey);

                using (var fs = new FileStream(savePath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                {
                    await file.CopyToAsync(fs);
                }

                string? sha256Hex = null;
                try
                {
                    using var sha = SHA256.Create();
                    await using var rs = new FileStream(savePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true);
                    var hash = await sha.ComputeHashAsync(rs);
                    sha256Hex = Convert.ToHexString(hash).ToLowerInvariant();
                }
                catch
                {
                    sha256Hex = null;
                }

                var contentType = string.IsNullOrWhiteSpace(file.ContentType) ? "application/octet-stream" : file.ContentType;
                var storagePath = Path.Combine(dayFolderName, fileKey).Replace('\\', '/');

                if (hasDb)
                {
                    var uploaderId = User.FindFirstValue(ClaimTypes.NameIdentifier);

                    await using var conn = new SqlConnection(cs);
                    await conn.OpenAsync();

                    await using var cmd = conn.CreateCommand();
                    cmd.CommandText = @"
INSERT INTO dbo.DocumentFiles
    (DocId, CommentId, FileKey, OriginalName, StoragePath, ContentType, ByteSize, Sha256, UploadedBy, UploadedAt)
VALUES
    (@DocId, @CommentId, @FileKey, @OriginalName, @StoragePath, @ContentType, @ByteSize, @Sha256, @UploadedBy, SYSUTCDATETIME());";

                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId! });
                    cmd.Parameters.Add(new SqlParameter("@CommentId", SqlDbType.BigInt) { Value = (object?)commentId ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@FileKey", SqlDbType.NVarChar, 200) { Value = fileKey });
                    cmd.Parameters.Add(new SqlParameter("@OriginalName", SqlDbType.NVarChar, 260) { Value = originalName });
                    cmd.Parameters.Add(new SqlParameter("@StoragePath", SqlDbType.NVarChar, 400) { Value = storagePath });
                    cmd.Parameters.Add(new SqlParameter("@ContentType", SqlDbType.NVarChar, 200) { Value = contentType });
                    cmd.Parameters.Add(new SqlParameter("@ByteSize", SqlDbType.BigInt) { Value = file.Length });
                    cmd.Parameters.Add(new SqlParameter("@Sha256", SqlDbType.NVarChar, 64) { Value = (object?)sha256Hex ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@UploadedBy", SqlDbType.NVarChar, 450) { Value = (object?)uploaderId ?? DBNull.Value });

                    await cmd.ExecuteNonQueryAsync();
                }

                items.Add(new UploadItemResult
                {
                    fileKey = fileKey,
                    originalName = originalName,
                    contentType = contentType,
                    byteSize = file.Length
                });
            }

            var errs = items.Where(x => x.error != null).ToArray();
            var oks = items.Where(x => x.error == null).ToArray();

            if (oks.Length == 0)
                return BadRequest(new { messages = errs.Select(e => e.error!).Distinct().ToArray() });

            return Json(new { items = oks, errors = errs });
        }

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
            else if (ext is ".gif") contentType = "image/gif";
            else if (ext is ".bmp") contentType = "image/bmp";
            else if (ext is ".webp") contentType = "image/webp";
            else if (ext is ".tif" or ".tiff") contentType = "image/tiff";
            else if (ext is ".svg") contentType = "image/svg+xml";
            else if (ext is ".ico") contentType = "image/x-icon";
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
