// 2026.01.23 Added: DocController 분리용 스켈레톤 생성

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Lib.Net.Http.WebPush;
using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services;
using static WebApplication1.Controllers.DocController;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class ComposeController : Controller
    {
        private readonly IStringLocalizer _S;
        private readonly IConfiguration _cfg;
        private readonly IAntiforgery _antiforgery;
        private readonly IWebHostEnvironment _env;
        private readonly ApplicationDbContext _db;
        private readonly IDocTemplateService _tpl;
        private readonly ILogger _log;
        private readonly IEmailSender _emailSender;
        private readonly SmtpOptions _smtpOpt;
        private static readonly object _attachSeqLock = new object();
        private readonly IWebPushNotifier _webPushNotifier;

        public ComposeController(
            IStringLocalizer S,
            IConfiguration cfg,
            IAntiforgery antiforgery,
            IWebHostEnvironment env,
            ApplicationDbContext db,
            IDocTemplateService tpl,
            IEmailSender emailSender,
            IOptions<SmtpOptions> smtpOptions,
            ILogger<ComposeController> log,
            IWebPushNotifier webPushNotifier
        )
        {
            _S = S;
            _cfg = cfg;
            _antiforgery = antiforgery;
            _env = env;
            _db = db;
            _tpl = tpl;
            _emailSender = emailSender;
            _smtpOpt = smtpOptions?.Value ?? new SmtpOptions();
            _log = log;
            _webPushNotifier = webPushNotifier;
        }

        [HttpGet("New")]
        public IActionResult New()        
        {
            var vm = new DocTLViewModel();

            var userComp = User.FindFirstValue("compCd") ?? "";
            var userDept = User.FindFirstValue("departmentId") ?? "";

            if (string.IsNullOrWhiteSpace(userComp) || string.IsNullOrWhiteSpace(userDept))
            {
                try
                {
                    var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
                    if (!string.IsNullOrEmpty(userId))
                    {
                        var cs = _cfg.GetConnectionString("DefaultConnection");
                        using var conn = new SqlConnection(cs);
                        using var cmd = new SqlCommand(
                            @"SELECT TOP 1 CompCd, DepartmentId
                              FROM dbo.UserProfiles
                              WHERE UserId = @userId", conn);
                        cmd.Parameters.Add(new SqlParameter("@userId", SqlDbType.NVarChar, 450) { Value = userId });
                        conn.Open();
                        using var rd = cmd.ExecuteReader();
                        if (rd.Read())
                        {
                            if (string.IsNullOrWhiteSpace(userComp) && !rd.IsDBNull(0))
                                userComp = rd.GetString(0) ?? "";
                            if (string.IsNullOrWhiteSpace(userDept) && !rd.IsDBNull(1))
                                userDept = rd.GetInt32(1).ToString();
                        }
                    }
                }
                catch { }
            }

            vm.CompOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "", Text = _S["_CM_Select"].Value, Selected = true }
            };

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                using var cmd = new SqlCommand("SELECT CompCd, Name FROM dbo.CompMasters ORDER BY CompCd", conn);
                conn.Open();
                using var rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    var compCd = rd.GetString(0);
                    var name = rd.IsDBNull(1) ? compCd : rd.GetString(1);
                    vm.CompOptions.Add(new SelectListItem
                    {
                        Value = compCd,
                        Text = name,
                        Selected = !string.IsNullOrWhiteSpace(userComp)
                                   && string.Equals(userComp, compCd, StringComparison.OrdinalIgnoreCase)
                    });
                }
                if (vm.CompOptions.Any(o => o.Selected))
                    vm.CompOptions[0].Selected = false;
            }
            catch { }

            vm.DepartmentOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "__SELECT__", Text = $"-- {_S["_CM_Select"]} --", Selected = true },
                new SelectListItem { Value = "", Text = _S["_CM_Common"].Value, Selected = string.IsNullOrEmpty(userDept) }
            };

            ViewBag.Templates = Array.Empty<(string code, string title)>();
            ViewBag.Kinds = Array.Empty<(string value, string text)>();
            ViewBag.Departments = Array.Empty<(string value, string text)>();
            ViewBag.Sites = Array.Empty<(string value, string text)>();
            ViewBag.UserComp = userComp;

            return View("Select", vm);
        }


        [HttpGet("Create")]
        public async Task<IActionResult> Create(string templateCode)
        {
            var meta = await _tpl.LoadMetaAsync(templateCode);

            var descriptorJson = string.IsNullOrWhiteSpace(meta.descriptorJson)
                ? "{}"
                : meta.descriptorJson;

            var previewJson = "{}";
            var excelAbsPath = meta.excelFilePath ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(excelAbsPath) && System.IO.File.Exists(excelAbsPath))
            {
                try
                {
                    var rebuilt = BuildPreviewJsonFromExcel(excelAbsPath);
                    if (!string.IsNullOrWhiteSpace(rebuilt) && HasCells(rebuilt))
                    {
                        previewJson = rebuilt;
                    }
                }
                catch
                {
                }
            }

            if (previewJson == "{}")
            {
                var metaPreview = string.IsNullOrWhiteSpace(meta.previewJson)
                    ? "{}"
                    : meta.previewJson;

                if (!string.IsNullOrWhiteSpace(metaPreview) && HasCells(metaPreview))
                {
                    previewJson = metaPreview;
                }
            }

            // ===== 기존 ADO 로더 블록 전체 제거하고 아래로 대체 =====
            try
            {
                var langCode = "ko";

                // (기존, 빌드 에러) var orgNodes = await BuildOrgTreeNodesAsync(conn, langCode);
                // (수정) BuildOrgTreeNodesAsync의 기존 시그니처에 맞춰 호출
                var orgNodes = await BuildOrgTreeNodesAsync(langCode);

                ViewBag.OrgTreeNodes = orgNodes;
            }
            catch (Exception ex)
            {
                ViewBag.OrgTreeNodes = Array.Empty<OrgTreeNode>();
            }

            descriptorJson = BuildDescriptorJsonWithFlowGroups(templateCode, descriptorJson);

            ViewBag.DescriptorJson = descriptorJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.TemplateTitle = meta.templateTitle ?? string.Empty;
            ViewBag.TemplateCode = templateCode;
            ViewBag.HideCellPicker = true;

            return View("Compose");

            static bool HasCells(string json)
            {
                if (string.IsNullOrWhiteSpace(json)) return false;
                if (!json.Contains("\"cells\"", StringComparison.Ordinal)) return false;
                try
                {
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("cells", out var cells)
                        && cells.ValueKind == JsonValueKind.Array
                        && cells.GetArrayLength() > 0)
                    {
                        var firstRow = cells[0];
                        return firstRow.ValueKind == JsonValueKind.Array;
                    }
                }
                catch { return false; }
                return false;
            }
        }
        // 2025.12.29 Changed: 문서 Create 성공 후 반환된 docId로만 업로드를 허용하고 Upload는 docId 필수와 CSRF 검증을 적용하여 문서 저장 후 첨부 업로드 정책을 강제 적용


        [HttpGet("Csrf")]
        [Produces("application/json")]
        public IActionResult Csrf()
        // ========== CSRF ==========
        {
            var tokens = _antiforgery.GetAndStoreTokens(HttpContext);
            return Json(new { headerName = "RequestVerificationToken", token = tokens.RequestToken });
        }


        [HttpPost("Upload")]
        [RequestSizeLimit(50_000_000)]
        public async Task<IActionResult> Upload([FromQuery] string? docId, [FromForm] List<IFormFile> files)
        {
            if (files == null || files.Count == 0)
                return BadRequest(new { messages = new[] { "DOC_Err_UploadFailed" }, detail = "no files" });

            var trimmedDocId = string.IsNullOrWhiteSpace(docId) ? null : docId.Trim();
            var uploadedBy = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

            // 정책: docId 없는 업로드는 허용하지 않음
            if (string.IsNullOrWhiteSpace(trimmedDocId))
                return BadRequest(new { messages = new[] { "DOC_Err_UploadFailed" }, detail = "docId required" });

            // 1) 문서 메타(CompCd/DepartCD/작성월) 해석
            var meta = await ResolveDocMetaForAttachmentPathAsync(trimmedDocId!);
            var compCd = string.IsNullOrWhiteSpace(meta.compCd) ? "0000" : meta.compCd!;
            var departCd = string.IsNullOrWhiteSpace(meta.departCd) ? "0" : meta.departCd!;
            var y = meta.yyyy;
            var m = meta.mm;

            // 2) 최종 저장 상대/절대 루트
            var relDir = Path.Combine("App_Data", "Docs", compCd, y, m, departCd);
            var absDir = Path.Combine(_env.ContentRootPath, relDir);
            Directory.CreateDirectory(absDir);

            _log.LogInformation("Doc.Upload attach start docId={docId} dir={dir} fileCount={cnt}", trimmedDocId, relDir, files.Count);

            var results = new List<object>();

            foreach (var f in files)
            {
                if (f == null || f.Length <= 0) continue;

                var originalName = Path.GetFileName(f.FileName);
                var ext = Path.GetExtension(originalName)?.ToLowerInvariant() ?? "";
                if (string.IsNullOrWhiteSpace(ext)) ext = "";

                var fileKey = Guid.NewGuid().ToString("N");

                // SHA256: DB varbinary(32)
                byte[] shaBytes;
                await using (var inStream = f.OpenReadStream())
                {
                    using var hasher = SHA256.Create();
                    shaBytes = await hasher.ComputeHashAsync(inStream);
                }

                // 3) Attach{N} 번호 결정 (DocId 기준 다음 번호)
                int nextN;
                lock (_attachSeqLock)
                {
                    // lock 범위 내에서 "다음 번호" 계산만 보호 (동시 업로드 최소 방어)
                    nextN = GetNextAttachSeqForDoc(trimmedDocId!);
                }

                var safeFileName = $"{trimmedDocId}_Attach{nextN}{ext}";
                var relPath = Path.Combine(relDir, safeFileName);      // DB 저장용(상대경로)
                var absPath = Path.Combine(absDir, safeFileName);      // 파일 저장용(절대경로)

                // 파일 저장
                await using (var outStream = System.IO.File.Create(absPath))
                {
                    await f.CopyToAsync(outStream);
                }

                var contentType = string.IsNullOrWhiteSpace(f.ContentType) ? "application/octet-stream" : f.ContentType;

                // 4) DB 저장: StoragePath는 상대경로만 저장
                try
                {
                    await InsertDocumentFileRowAsync(
                        trimmedDocId!,
                        fileKey,
                        originalName,
                        relPath,          // ★ 상대경로 저장
                        contentType,
                        (long)f.Length,
                        shaBytes,         // varbinary(32)
                        uploadedBy
                    );
                }
                catch (SqlException se)
                {
                    _log.LogError(se, "Doc.Upload SQL failed docId={docId} fileKey={fileKey}", trimmedDocId, fileKey);
                    return StatusCode(500, new { messages = new[] { "DOC_Err_UploadFailed" }, detail = se.Message });
                }
                catch (Exception ex)
                {
                    _log.LogError(ex, "Doc.Upload insert failed docId={docId} fileKey={fileKey}", trimmedDocId, fileKey);
                    return StatusCode(500, new { messages = new[] { "DOC_Err_UploadFailed" }, detail = ex.Message });
                }

                results.Add(new
                {
                    fileKey,
                    originalName,
                    contentType,
                    byteSize = (long)f.Length,
                    docId = trimmedDocId,
                    storagePath = relPath.Replace('/', '\\') // UI/디버그 확인용(표시)
                });
            }

            _log.LogInformation("Doc.Upload attach done docId={docId} savedCount={cnt}", trimmedDocId, results.Count);
            return Ok(new { files = results });
        }


        private async Task InsertDocumentFileRowAsync(string docId, string fileKey, string originalName, string storagePath, string contentType, long byteSize, byte[] sha256Bytes, string uploadedBy)
        {
            if (sha256Bytes == null || sha256Bytes.Length != 32)
                throw new InvalidOperationException($"Sha256Bytes must be 32 bytes. actual={(sha256Bytes == null ? 0 : sha256Bytes.Length)}");

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            await using var cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = @"
INSERT INTO dbo.DocumentFiles
(DocId, FileKey, OriginalName, StoragePath, ContentType, ByteSize, Sha256, UploadedBy, UploadedAt)
VALUES
(@DocId, @FileKey, @OriginalName, @StoragePath, @ContentType, @ByteSize, @Sha256, @UploadedBy, SYSUTCDATETIME());";

            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 100) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@FileKey", SqlDbType.NVarChar, 64) { Value = fileKey });
            cmd.Parameters.Add(new SqlParameter("@OriginalName", SqlDbType.NVarChar, 510) { Value = originalName ?? "" });
            cmd.Parameters.Add(new SqlParameter("@StoragePath", SqlDbType.NVarChar, 800) { Value = storagePath ?? "" }); // ★ 상대경로
            cmd.Parameters.Add(new SqlParameter("@ContentType", SqlDbType.NVarChar, 254) { Value = contentType ?? "" });
            cmd.Parameters.Add(new SqlParameter("@ByteSize", SqlDbType.BigInt) { Value = byteSize });

            cmd.Parameters.Add(new SqlParameter("@Sha256", SqlDbType.VarBinary, 32) { Value = sha256Bytes });
            cmd.Parameters.Add(new SqlParameter("@UploadedBy", SqlDbType.NVarChar, 128) { Value = uploadedBy ?? "" });

            await cmd.ExecuteNonQueryAsync();
        }

        private async Task<(string? compCd, string? departCd, string yyyy, string mm)> ResolveDocMetaForAttachmentPathAsync(string docId)
        {
            var now = DateTime.UtcNow;
            var yyyy = now.ToString("yyyy");
            var mm = now.ToString("MM");

            string? compCd = null;
            string? departCd = null;

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"
SELECT TOP (1)
    CompCd,
    DepartmentId
FROM dbo.Documents
WHERE DocId = @DocId;";

                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 100) { Value = docId });

                await using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    compCd = r["CompCd"] as string;
                    departCd = (r["DepartmentId"] == DBNull.Value ? null : Convert.ToString(r["DepartmentId"]));
                }
            }
            catch
            {
                // 실패 시 기본값(0000/0)으로 상위에서 보정
            }

            return (compCd, departCd, yyyy, mm);
        }

        // DocId 기준 Attach{N} 다음 번호: StoragePath에 "_Attach{N}" 패턴이 있다고 가정하고 MAX+1
        private int GetNextAttachSeqForDoc(string docId)        
        {
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                if (string.IsNullOrWhiteSpace(cs)) return 1;

                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"
SELECT MAX(TRY_CONVERT(int,
    SUBSTRING(StoragePath,
              CHARINDEX('_Attach', StoragePath) + LEN('_Attach'),
              CHARINDEX('.', StoragePath, CHARINDEX('_Attach', StoragePath)) - (CHARINDEX('_Attach', StoragePath) + LEN('_Attach'))
    )))
FROM dbo.DocumentFiles
WHERE DocId = @DocId
  AND StoragePath LIKE '%\_Attach%.%' ESCAPE '\';";

                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 100) { Value = docId });

                var obj = cmd.ExecuteScalar();
                if (obj == null || obj == DBNull.Value) return 1;

                var maxN = Convert.ToInt32(obj);
                if (maxN < 1) return 1;
                return maxN + 1;
            }
            catch
            {
                return 1;
            }
        }

        [HttpPost("Create")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> Create([FromBody] ComposePostDto dto)
        {
            // ------------------------------------------------------------
            // 방법 B 문서 저장 먼저 첨부는 성공 후 별도 업로드 전제
            // ------------------------------------------------------------

            ComposePostDto? resolvedDto = dto;

            // 415 방지 form-data 로 들어오면 JSON payload만 추출해서 dto로 복원
            if ((resolvedDto == null || string.IsNullOrWhiteSpace(resolvedDto.TemplateCode)) && Request.HasFormContentType)
            {
                try
                {
                    var form = await Request.ReadFormAsync();

                    var json =
                        form["payload"].FirstOrDefault()
                        ?? form["dto"].FirstOrDefault()
                        ?? form["data"].FirstOrDefault();

                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        resolvedDto = System.Text.Json.JsonSerializer.Deserialize<ComposePostDto>(
                            json,
                            new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true }
                        );
                    }

                    // 파일은 방법 B 정책상 Create에서 절대 처리하지 않음
                    // var ignoredFiles = form.Files;
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "Create: form-data dto parse failed");
                }
            }

            dto = resolvedDto!;

            if (dto is null || string.IsNullOrWhiteSpace(dto.TemplateCode))
                return BadRequest(new { messages = new[] { "DOC_Val_TemplateRequired" }, stage = "arg", detail = "templateCode null/empty" });

            var tc = dto.TemplateCode!;
            var inputsMap = dto.Inputs ?? new Dictionary<string, string>();

            var userCompDept = GetUserCompDept();
            var compCd = string.IsNullOrWhiteSpace(userCompDept.compCd) ? "0000" : userCompDept.compCd!;
            var deptIdPart = string.IsNullOrWhiteSpace(userCompDept.departmentId) ? "0" : userCompDept.departmentId!;

            var (descriptorJson, _previewJsonFromTpl, title, versionId, excelPathRaw) = await _tpl.LoadMetaAsync(tc);

            // ✅ 방어: DB에 개발환경 절대경로가 섞여도 App_Data 이후 상대경로로 정규화
            excelPathRaw = NormalizeTemplateExcelPath(excelPathRaw);

            var excelPath = ToContentRootAbsolute(excelPathRaw);

            if (versionId <= 0 || string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_TemplateNotReady" },
                    stage = "generate",
                    detail = $"Excel not found or version invalid. versionId={versionId}, path='{excelPathRaw}' → '{excelPath}'"
                });

            string tempExcelFullPath;
            try
            {
                _log.LogInformation("GenerateExcel start tc={tc} ver={ver} path={path} inputs={cnt}", tc, versionId, excelPath, inputsMap.Count);

                tempExcelFullPath = await GenerateExcelFromInputsAsync(
                    versionId, excelPath, inputsMap, dto.DescriptorVersion);

                _log.LogInformation("GenerateExcel done out={out}", tempExcelFullPath);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Create.Generate failed tc={tc} ver={ver} path={path}", tc, versionId, excelPath);
                return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "generate", detail = ex.Message });
            }

            var now = DateTime.Now;
            var year = now.Year.ToString("D4");
            var month = now.Month.ToString("D2");

            var destDirRel = Path.Combine("App_Data", "Docs", compCd, year, month, deptIdPart);

            var docId = Path.GetFileNameWithoutExtension(tempExcelFullPath);
            var excelExt = Path.GetExtension(tempExcelFullPath);
            if (string.IsNullOrWhiteSpace(excelExt)) excelExt = ".xlsx";

            var outputPathForDb = Path.Combine(destDirRel, $"{docId}{excelExt}");

            var filledPreviewJson = BuildPreviewJsonFromExcel(tempExcelFullPath);
            var normalizedDesc = BuildDescriptorJsonWithFlowGroups(tc, descriptorJson);

            // ----------------------------
            // 1) 수신자 먼저 해석 toEmails 확보
            // ----------------------------
            List<string> toEmails;
            List<string> diag;
            try
            {
                var recipients = GetInitialRecipients(dto, normalizedDesc, _cfg, docId, tc);
                toEmails = recipients.emails ?? new List<string>();
                diag = recipients.diag ?? new List<string>();
            }
            catch
            {
                toEmails = new();
                diag = new() { "recipients resolve error" };
            }

            // ----------------------------
            // 2) 결재 JSON 추출 및 보정
            // ----------------------------
            var approvalsJson = ExtractApprovalsJson(normalizedDesc);

            bool approvalsHasAnyValue = false;
            bool approvalsParsedOk = false;

            if (!string.IsNullOrWhiteSpace(approvalsJson))
            {
                try
                {
                    using var aj = System.Text.Json.JsonDocument.Parse(approvalsJson);
                    if (aj.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        approvalsParsedOk = true;
                        foreach (var a in aj.RootElement.EnumerateArray())
                        {
                            string? v = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == System.Text.Json.JsonValueKind.String) v = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == System.Text.Json.JsonValueKind.String) v = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == System.Text.Json.JsonValueKind.String) v = v3.GetString();

                            if (!string.IsNullOrWhiteSpace(v))
                            {
                                approvalsHasAnyValue = true;
                                break;
                            }
                        }
                    }
                }
                catch
                {
                    approvalsParsedOk = false;
                    approvalsHasAnyValue = false;
                }
            }

            if (string.IsNullOrWhiteSpace(approvalsJson)
                || approvalsJson.Trim() == "[]"
                || approvalsParsedOk == false
                || approvalsHasAnyValue == false)
            {
                if (toEmails.Count > 0)
                {
                    var built = new List<Dictionary<string, object>>();

                    int seq = 1;
                    foreach (var raw in toEmails)
                    {
                        var email = (raw ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(email)) continue;

                        built.Add(new Dictionary<string, object>
                        {
                            ["roleKey"] = $"A{seq}",
                            ["approverType"] = "Person",
                            ["required"] = true,
                            ["value"] = email
                        });

                        seq++;
                    }

                    if (built.Count == 0)
                        return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "approvals", detail = "No valid approver emails." });

                    approvalsJson = System.Text.Json.JsonSerializer.Serialize(built);
                }
                else
                {
                    return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "approvals", detail = "Approver not resolved (toEmails empty)." });
                }
            }
            else
            {
                try
                {
                    using var aj = System.Text.Json.JsonDocument.Parse(approvalsJson);
                    if (aj.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        var list = new List<Dictionary<string, object>>();
                        int seq = 1;

                        foreach (var a in aj.RootElement.EnumerateArray())
                        {
                            var item = new Dictionary<string, object>();

                            string at = "Person";
                            if (a.TryGetProperty("approverType", out var at1) && at1.ValueKind == System.Text.Json.JsonValueKind.String)
                                at = string.IsNullOrWhiteSpace(at1.GetString()) ? at : at1.GetString()!;
                            else if (a.TryGetProperty("ApproverType", out var at2) && at2.ValueKind == System.Text.Json.JsonValueKind.String)
                                at = string.IsNullOrWhiteSpace(at2.GetString()) ? at : at2.GetString()!;

                            bool required = false;
                            if (a.TryGetProperty("required", out var r1) && (r1.ValueKind == System.Text.Json.JsonValueKind.True || r1.ValueKind == System.Text.Json.JsonValueKind.False))
                                required = r1.GetBoolean();

                            string? v = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == System.Text.Json.JsonValueKind.String) v = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == System.Text.Json.JsonValueKind.String) v = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == System.Text.Json.JsonValueKind.String) v = v3.GetString();

                            item["roleKey"] = $"A{seq}";
                            item["approverType"] = at;
                            item["required"] = required;
                            item["value"] = v ?? string.Empty;

                            list.Add(item);
                            seq++;
                        }

                        approvalsJson = System.Text.Json.JsonSerializer.Serialize(list);
                    }
                }
                catch
                {
                    // 정규화 실패는 원본 유지
                }
            }

            try
            {
                var pair = await SaveDocumentWithCollisionGuardAsync(
                    docId: docId,
                    templateCode: tc,
                    title: string.IsNullOrWhiteSpace(title) ? tc : title!,
                    status: "PendingA1",
                    outputPath: outputPathForDb,
                    inputs: inputsMap,
                    approvalsJson: approvalsJson,
                    descriptorJson: normalizedDesc,
                    userId: User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "",
                    userName: User?.Identity?.Name ?? "",
                    compCd: compCd,
                    departmentId: deptIdPart
                );

                docId = pair.docId;
                outputPathForDb = pair.outputPath;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Create.DB save failed with collision guard docId={docId} tc={tc}", docId, tc);

                object? sqlDiag = null;
                if (ex is SqlException se) sqlDiag = new { se.Number, se.State, se.Procedure, se.LineNumber };

                return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "db", detail = ex.Message, sql = sqlDiag });
            }

            string finalExcelFullPath = string.Empty;
            try
            {
                finalExcelFullPath = ToContentRootAbsolute(outputPathForDb);

                var finalDir = Path.GetDirectoryName(finalExcelFullPath);
                if (!string.IsNullOrWhiteSpace(finalDir))
                    Directory.CreateDirectory(finalDir);

                if (!string.Equals(tempExcelFullPath, finalExcelFullPath, StringComparison.OrdinalIgnoreCase))
                {
                    if (!System.IO.File.Exists(finalExcelFullPath))
                    {
                        System.IO.File.Move(tempExcelFullPath, finalExcelFullPath, overwrite: false);
                        tempExcelFullPath = finalExcelFullPath;
                    }
                    else
                    {
                        _log.LogWarning("Final excel already exists. skip move. docId={docId} final={final}", docId, finalExcelFullPath);
                    }
                }

                if (!string.IsNullOrWhiteSpace(finalExcelFullPath) && System.IO.File.Exists(finalExcelFullPath))
                    filledPreviewJson = BuildPreviewJsonFromExcel(finalExcelFullPath);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Move generated excel failed temp={temp} final={final} docId={docId}",
                    tempExcelFullPath, finalExcelFullPath, docId);
            }

            try
            {
                await EnsureApprovalsAndSyncAsync(docId, tc);
                await FillDocumentApprovalsFromEmailsAsync(docId, toEmails);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "EnsureApprovalsAndSync failed docId={docId} tc={tc}", docId, tc);
            }

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection");
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();
                await using var fix = conn.CreateCommand();
                fix.CommandText = @"
UPDATE a
   SET a.UserId = u.Id
FROM dbo.DocumentApprovals a
JOIN dbo.AspNetUsers u
  ON (a.ApproverValue = u.Email OR a.ApproverValue = u.UserName OR a.ApproverValue = u.Id)
WHERE a.DocId = @DocId
  AND a.UserId IS NULL
  AND a.ApproverValue IS NOT NULL AND LTRIM(RTRIM(a.ApproverValue)) <> N'';";
                fix.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                await fix.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Post-fix UserId update failed docId={docId}", docId);
            }

            // 공유 저장 로직은 기존 그대로 유지
            try
            {
                var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";

                var selected = (dto.SelectedRecipientUserIds ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (selected.Count > 0)
                {
                    var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                    await using var conn = new SqlConnection(cs);
                    await conn.OpenAsync();
                    await using var tx = await conn.BeginTransactionAsync();

                    try
                    {
                        foreach (var targetUserId in selected)
                        {
                            await using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = (SqlTransaction)tx;
                                cmd.CommandText = @"
IF EXISTS (SELECT 1 FROM dbo.DocumentShares WHERE DocId=@DocId AND UserId=@UserId)
BEGIN
    UPDATE dbo.DocumentShares
       SET IsRevoked = 0,
           ExpireAt = NULL
     WHERE DocId=@DocId AND UserId=@UserId;
END
ELSE
BEGIN
    INSERT INTO dbo.DocumentShares (DocId, UserId, AccessRole, ExpireAt, IsRevoked, CreatedBy, CreatedAt)
    VALUES (@DocId, @UserId, 'Commenter', NULL, 0, @CreatedBy, SYSUTCDATETIME());
END";
                                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                                cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 64) { Value = actorId });
                                await cmd.ExecuteNonQueryAsync();
                            }

                            await using (var logCmd = conn.CreateCommand())
                            {
                                logCmd.Transaction = (SqlTransaction)tx;
                                logCmd.CommandText = @"
INSERT INTO dbo.DocumentShareLogs (DocId, ActorId, ChangeCode, TargetUserId, BeforeJson, AfterJson, ChangedAt)
VALUES (@DocId, @ActorId, @ChangeCode, @TargetUserId, NULL, @AfterJson, SYSUTCDATETIME());";
                                logCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                                logCmd.Parameters.Add(new SqlParameter("@ActorId", SqlDbType.NVarChar, 64) { Value = actorId });
                                logCmd.Parameters.Add(new SqlParameter("@ChangeCode", SqlDbType.NVarChar, 50) { Value = "ShareAdded" });
                                logCmd.Parameters.Add(new SqlParameter("@TargetUserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                                logCmd.Parameters.Add(new SqlParameter("@AfterJson", SqlDbType.NVarChar, -1) { Value = "{\"accessRole\":\"Commenter\"}" });
                                await logCmd.ExecuteNonQueryAsync();
                            }
                        }

                        await ((SqlTransaction)tx).CommitAsync();
                    }
                    catch
                    {
                        await ((SqlTransaction)tx).RollbackAsync();
                        throw;
                    }

                    _log.LogInformation("Shares saved docId={docId} count={cnt}", docId, selected.Count);
                }
                else
                {
                    _log.LogInformation("No shares selected docId={docId}", docId);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Share save failed docId={docId}", docId);
            }

            // ------------------------------------------------------------
            // ✅ WebPush: A1(첫 결재자) + 공유자에게 "카운트" 알림 발송 (URL 없음, tag 고정)
            // - 커밋/저장 완료 이후에만 발송 (현재 위치 OK)
            // - 중복 알림 폭주 방지: tag 고정
            //   결재: badge-approval-pending
            //   공유: badge-shared
            // ------------------------------------------------------------
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";

                async Task<int> GetApprovalPendingCountAsync(string targetUserId)
                {
                    await using var conn = new SqlConnection(cs);
                    await conn.OpenAsync();

                    var sql = @"
SELECT COUNT(1)
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
WHERE a.UserId = @UserId
  AND ISNULL(a.Status, N'') = N'Pending'
  AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10))
  AND ISNULL(d.Status, N'') NOT LIKE N'Recalled%';";

                    await using var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                    return Convert.ToInt32(await cmd.ExecuteScalarAsync());
                }

                async Task<int> GetSharedUnreadCountAsync(string targetUserId)
                {
                    await using var conn = new SqlConnection(cs);
                    await conn.OpenAsync();

                    var sql = @"
SELECT COUNT(1)
FROM DocumentShares s
WHERE s.UserId = @UserId
  AND ISNULL(s.IsRevoked, 0) = 0
  AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME())
  AND NOT EXISTS (
      SELECT 1
      FROM DocumentViewLogs v
      WHERE v.DocId = s.DocId
        AND v.ViewerId = @UserId
        AND ISNULL(v.ViewerRole, N'') = N'Shared'
  );";

                    await using var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                    return Convert.ToInt32(await cmd.ExecuteScalarAsync());
                }

                async Task<string> GetA1ApproverUserIdByDocAsync(string xDocId)
                {
                    // EnsureApprovalsAndSync + UserId 보정까지 끝난 이후라면
                    // DocumentApprovals에서 StepOrder=1 UserId를 직접 읽는 게 가장 안전합니다.
                    await using var conn = new SqlConnection(cs);
                    await conn.OpenAsync();

                    var sql = @"
SELECT TOP (1) a.UserId
FROM dbo.DocumentApprovals a
WHERE a.DocId = @DocId
  AND a.StepOrder = 1
ORDER BY a.Id;";

                    await using var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = xDocId });

                    var obj = await cmd.ExecuteScalarAsync();
                    return (obj == null || obj == DBNull.Value) ? "" : (obj.ToString() ?? "").Trim();
                }

                // 1) A1 결재자에게 "결재 대기 N건" (tag 고정, URL 없음)
                var a1UserId = await GetA1ApproverUserIdByDocAsync(docId);
                if (!string.IsNullOrWhiteSpace(a1UserId))
                {
                    var n = await GetApprovalPendingCountAsync(a1UserId);

                    await _webPushNotifier.SendToUserIdAsync(
                        userId: a1UserId,
                        title: "E-BOARD",
                        body: $"{n}개의 문서가 결재 대기 중입니다.",
                        url: "/",                 // URL 사용 안 함(기본값)
                        tag: "badge-approval-pending"
                    );
                }
                else
                {
                    _log.LogWarning("WebPush A1 userId not found docId={docId}", docId);
                }

                // 2) 공유자에게 "공유 중(미확인) N건" (tag 고정, URL 없음)
                var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
                var shareIds = (dto.SelectedRecipientUserIds ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var uid in shareIds)
                {
                    var n = await GetSharedUnreadCountAsync(uid);

                    await _webPushNotifier.SendToUserIdAsync(
                        userId: uid,
                        title: "E-BOARD",
                        body: $"{n}개의 문서가 공유 중입니다.",
                        url: "/",             // URL 사용 안 함(기본값)
                        tag: "badge-shared"
                    );
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "WebPush notify failed docId={docId}", docId);
            }

            object? mailInfo = null;
            var uploadUrl = $"/Doc/Upload?docId={Uri.EscapeDataString(docId)}";

            return Json(new
            {
                ok = true,
                docId,
                title = title ?? tc,
                status = "PendingA1",
                previewJson = filledPreviewJson,
                approvalsJson,
                attachments = Array.Empty<object>(),
                mailInfo = mailInfo,
                uploadUrl = uploadUrl
            });
        }

        
        private async Task<object?> SendSubmitMailAsync(string docId, string templateCode, string title, string authorName, List<string> toEmails, List<string>? ccEmails = null, List<string>? bccEmails = null, List<string>? diag = null)
        {
            if (dto?.Mail == null || dto.Mail.Send == false)
                return new
                {
                    from = "",
                    fromDisplay = "",
                    to = Array.Empty<string>(),
                    toDisplay = Array.Empty<string>(),
                    cc = Array.Empty<string>(),
                    ccDisplay = Array.Empty<string>(),
                    bcc = Array.Empty<string>(),
                    bccDisplay = Array.Empty<string>(),
                    subject = "",
                    body = "",
                    sent = false,
                    sentCount = 0,
                    error = ""
                };

            static List<string> ResolveFirstStepApproverEmails(string? desc, IEnumerable<string> fallback)
            {
                var result = new List<string>();
                var useFallback = false;

                try
                {
                    if (string.IsNullOrWhiteSpace(desc))
                    {
                        useFallback = true;
                    }
                    else
                    {
                        using var dj = JsonDocument.Parse(desc);
                        if (!dj.RootElement.TryGetProperty("approvals", out var arr) ||
                            arr.ValueKind != JsonValueKind.Array)
                        {
                            useFallback = true;
                        }
                        else
                        {
                            int minOrder = int.MaxValue;

                            foreach (var a in arr.EnumerateArray())
                            {
                                int ord = 1;
                                if (a.TryGetProperty("order", out var o))
                                {
                                    if (o.ValueKind == JsonValueKind.Number && o.TryGetInt32(out var oi))
                                        ord = oi;
                                    else if (o.ValueKind == JsonValueKind.String &&
                                             int.TryParse(o.GetString(), out var os))
                                        ord = os;
                                }
                                if (ord < minOrder) minOrder = ord;
                            }

                            if (minOrder == int.MaxValue)
                            {
                                useFallback = true;
                            }
                            else
                            {
                                void AddEmail(string? e)
                                {
                                    if (!string.IsNullOrWhiteSpace(e))
                                        result.Add(e.Trim());
                                }

                                foreach (var a in arr.EnumerateArray())
                                {
                                    int ord = 1;
                                    if (a.TryGetProperty("order", out var o))
                                    {
                                        if (o.ValueKind == JsonValueKind.Number && o.TryGetInt32(out var oi))
                                            ord = oi;
                                        else if (o.ValueKind == JsonValueKind.String &&
                                                 int.TryParse(o.GetString(), out var os))
                                            ord = os;
                                    }
                                    if (ord != minOrder) continue;

                                    if (a.TryGetProperty("users", out var users) && users.ValueKind == JsonValueKind.Array)
                                    {
                                        foreach (var u in users.EnumerateArray())
                                        {
                                            if (u.ValueKind != JsonValueKind.Object) continue;
                                            if (u.TryGetProperty("email", out var pe) && pe.ValueKind == JsonValueKind.String)
                                                AddEmail(pe.GetString());
                                            else if (u.TryGetProperty("mail", out var pm) && pm.ValueKind == JsonValueKind.String)
                                                AddEmail(pm.GetString());
                                            else if (u.TryGetProperty("Email", out var pE) && pE.ValueKind == JsonValueKind.String)
                                                AddEmail(pE.GetString());
                                        }
                                    }

                                    if (a.TryGetProperty("user", out var user) && user.ValueKind == JsonValueKind.Object)
                                    {
                                        if (user.TryGetProperty("email", out var pe) && pe.ValueKind == JsonValueKind.String)
                                            AddEmail(pe.GetString());
                                        else if (user.TryGetProperty("mail", out var pm) && pm.ValueKind == JsonValueKind.String)
                                            AddEmail(pm.GetString());
                                        else if (user.TryGetProperty("Email", out var pE) && pE.ValueKind == JsonValueKind.String)
                                            AddEmail(pE.GetString());
                                    }

                                    if (a.TryGetProperty("emails", out var emails) && emails.ValueKind == JsonValueKind.Array)
                                    {
                                        foreach (var e in emails.EnumerateArray())
                                            if (e.ValueKind == JsonValueKind.String)
                                                AddEmail(e.GetString());
                                    }

                                    if (a.TryGetProperty("email", out var email) && email.ValueKind == JsonValueKind.String)
                                        AddEmail(email.GetString());

                                    if (a.TryGetProperty("value", out var v) && v.ValueKind == JsonValueKind.String)
                                        AddEmail(v.GetString());
                                    if (a.TryGetProperty("ApproverValue", out var av) && av.ValueKind == JsonValueKind.String)
                                        AddEmail(av.GetString());
                                }
                            }
                        }
                    }
                }
                catch
                {
                    useFallback = true;
                }

                if (result.Count == 0 && useFallback && fallback != null)
                    result.AddRange(fallback);

                return result
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => s.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            IEnumerable<string> Norm(IEnumerable<string>? src) =>
                (src ?? Array.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s))
                                              .Select(s => s.Trim())
                                              .Distinct(StringComparer.OrdinalIgnoreCase);

            var fromEmail = !string.IsNullOrWhiteSpace(dto?.Mail?.From) ? dto!.Mail!.From! : GetCurrentUserEmail();
            var authorDisplay = GetCurrentUserDisplayNameStrict();
            var fromAddress = ComposeAddress(fromEmail, authorDisplay);

            var firstStepToEmails = ResolveFirstStepApproverEmails(normalizedDesc, toEmails);

            if (!Norm(firstStepToEmails).Any())
            {
                try
                {
                    var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                    await using var conn = new SqlConnection(cs);
                    await conn.OpenAsync();

                    await using var cmd = conn.CreateCommand();
                    cmd.CommandText = @"ㅊ
        SELECT DISTINCT
               COALESCE(u.Email, a.ApproverValue) AS Email
        FROM dbo.DocumentApprovals a
        LEFT JOIN dbo.AspNetUsers u ON a.UserId = u.Id
        WHERE a.DocId = @DocId
          AND a.Status = N'Pending'
          AND a.StepOrder = (
                SELECT MIN(StepOrder)
                FROM dbo.DocumentApprovals
                WHERE DocId = @DocId
                  AND Status = N'Pending'
              );";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                    var fallbackEmails = new List<string>();
                    await using (var rdr = await cmd.ExecuteReaderAsync())
                    {
                        while (await rdr.ReadAsync())
                        {
                            var e = rdr.IsDBNull(0) ? null : rdr.GetString(0);
                            if (!string.IsNullOrWhiteSpace(e))
                                fallbackEmails.Add(e.Trim());
                        }
                    }

                    if (fallbackEmails.Count > 0)
                        firstStepToEmails = fallbackEmails;
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "First step approver email fallback from DocumentApprovals failed docId={docId}", docId);
                }
            }

            var toList = firstStepToEmails;
            var ccList = dto?.Mail?.CC ?? new List<string>();
            var bccList = dto?.Mail?.BCC ?? new List<string>();

            List<string> toDisplay = Norm(toList).Select(e => ComposeAddress(e, GetDisplayNameByEmailStrict(e))).ToList();
            List<string> ccDisplay = Norm(ccList).Select(e => ComposeAddress(e, GetDisplayNameByEmailStrict(e))).ToList();
            List<string> bccDisplay = Norm(bccList).Select(e => ComposeAddress(e, GetDisplayNameByEmailStrict(e))).ToList();

            string subject, body;
            {
                var reqSubj = dto?.Mail?.Subject;
                var reqBody = dto?.Mail?.Body;

                if (!string.IsNullOrWhiteSpace(reqSubj) && !string.IsNullOrWhiteSpace(reqBody))
                {
                    subject = reqSubj!;
                    body = reqBody!;
                }
                else
                {
                    var firstEmail = Norm(toList).FirstOrDefault();
                    var recipientDisplay = GetDisplayNameByEmailStrict(firstEmail);
                    var docTitle = string.IsNullOrWhiteSpace(title) ? tc : title!;
                    var tplMail = BuildSubmissionMail(authorDisplay, docTitle, docId, recipientDisplay);
                    subject = string.IsNullOrWhiteSpace(reqSubj) ? tplMail.subject : reqSubj!;
                    body = string.IsNullOrWhiteSpace(reqBody) ? tplMail.bodyHtml : reqBody!;
                }
            }

            var shouldSend = dto?.Mail?.Send != false;
            int sentCount = 0; string? sendErr = null;

            if (shouldSend && toDisplay.Any())
            {
                try
                {
                    var allTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var sendTargets = new List<string>();

                    void AddTargets(IEnumerable<string> src)
                    {
                        foreach (var addr in src)
                        {
                            if (string.IsNullOrWhiteSpace(addr)) continue;
                            if (allTargets.Add(addr))
                                sendTargets.Add(addr);
                        }
                    }

                    AddTargets(toDisplay);
                    AddTargets(ccDisplay);
                    AddTargets(bccDisplay);

                    foreach (var target in sendTargets)
                        await _emailSender.SendEmailAsync(target, subject, body);

                    sentCount = sendTargets.Count;
                }
                catch (Exception ex)
                {
                    sendErr = ex.Message;
                    _log.LogError(ex, "Mail send failed docId={docId}", docId);
                }
            }

            return new
            {
                from = fromEmail,
                fromDisplay = fromAddress,
                to = Norm(toList).ToArray(),
                toDisplay = toDisplay.ToArray(),
                cc = Norm(ccList).ToArray(),
                ccDisplay = ccDisplay.ToArray(),
                bcc = Norm(bccList).ToArray(),
                bccDisplay = bccDisplay.ToArray(),
                subject = subject,
                body = body,
                sent = (sentCount > 0 && string.IsNullOrWhiteSpace(sendErr)),
                sentCount = sentCount,
                error = sendErr ?? ""
            };
        }

        private async Task NotifySharesByWebPushAsync(SqlConnection conn, string docId, List<string> shareUserIds)
        {
            if (conn == null) throw new ArgumentNullException(nameof(conn));
            if (string.IsNullOrWhiteSpace(docId)) return;
            if (shareUserIds == null || shareUserIds.Count == 0) return;

            var targets = shareUserIds
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (targets.Count == 0) return;

            var subject = _cfg["WebPush:Subject"] ?? "mailto:admin@example.com";
            var publicKey = _cfg["WebPush:PublicKey"];
            var privateKey = _cfg["WebPush:PrivateKey"];
            if (string.IsNullOrWhiteSpace(publicKey) || string.IsNullOrWhiteSpace(privateKey)) return;

            var vapid = new Lib.Net.Http.WebPush.Authentication.VapidDetails(subject, publicKey, privateKey);
            var client = new Lib.Net.Http.WebPush.WebPushClient();

            var paramNames = new List<string>();
            using var cmd = conn.CreateCommand();
            for (int i = 0; i < targets.Count; i++)
            {
                var pName = "@u" + i.ToString(CultureInfo.InvariantCulture);
                paramNames.Add(pName);
                cmd.Parameters.AddWithValue(pName, targets[i]);
            }

            cmd.CommandText = $@"
SELECT Id, UserId, Endpoint, P256dh, Auth
FROM dbo.WebPushSubscriptions WITH (NOLOCK)
WHERE IsActive = 1
  AND UserId IN ({string.Join(",", paramNames)})
";

            var subs = new List<(long Id, string UserId, string Endpoint, string P256dh, string Auth)>();
            using (var rd = await cmd.ExecuteReaderAsync())
            {
                while (await rd.ReadAsync())
                {
                    subs.Add((
                        rd.GetInt64(0),
                        rd.GetString(1),
                        rd.GetString(2),
                        rd.GetString(3),
                        rd.GetString(4)
                    ));
                }
            }

            if (subs.Count == 0) return;

            var payloadObj = new
            {
                title = "E-BOARD",
                body = "작성된 문서가 공유되었습니다.",
                tag = "badge-shared",
                docId = docId
            };
            var payload = System.Text.Json.JsonSerializer.Serialize(payloadObj);

            foreach (var s in subs)
            {
                var pushSub = new Lib.Net.Http.WebPush.PushSubscription(
                    s.Endpoint,
                    s.P256dh,
                    s.Auth
                );

                try
                {
                    await client.SendNotificationAsync(pushSub, payload, vapid);
                }
                catch (Lib.Net.Http.WebPush.WebPushException)
                {
                    // 실패 처리(선택)
                }
                catch
                {
                    // 실패 처리(선택)
                }
            }
        }

        /// 현재 로그인 사용자가 주어진 문서에 대해  회수 가능 여부(CanRecall) 승인/보류/반려 가능 여부(CanApprove)를 판단.
        // 2025.11.20 Changed: GetApprovalCapabilitiesAsync 에 Hold/Reject 권한 플래그 추가 (튜플 2개 → 4개로 확장, Detail 뷰와 정합)
        // ========= 입력값 → 엑셀 =========
        private async Task<string> GenerateExcelFromInputsAsync(long templateVersionId, string templateExcelFullPath, Dictionary<string, string> inputs, string? descriptorVersion)
        {
            if (string.IsNullOrWhiteSpace(templateExcelPath) || !System.IO.File.Exists(templateExcelPath))
                throw new FileNotFoundException("Template excel not found", templateExcelPath);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var maps = new List<(string key, string a1, string type)>();

            await using (var conn = new SqlConnection(cs))
            {
                await conn.OpenAsync();
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
SELECT [Key], [Type], A1, CellRow, CellColumn
FROM dbo.DocTemplateField
WHERE VersionId = @vid
ORDER BY Id;";
                    cmd.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = versionId });

                    await using var rd = await cmd.ExecuteReaderAsync();
                    while (await rd.ReadAsync())
                    {
                        var key = rd["Key"] as string ?? "";
                        var typ = rd["Type"] as string ?? "Text";
                        string a1 = rd["A1"] as string ?? "";
                        if (string.IsNullOrWhiteSpace(a1))
                        {
                            int r = rd["CellRow"] is DBNull ? 0 : Convert.ToInt32(rd["CellRow"]);
                            int c = rd["CellColumn"] is DBNull ? 0 : Convert.ToInt32(rd["CellColumn"]);
                            if (r > 0 && c > 0) a1 = A1FromRowCol(r, c);
                        }
                        if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(a1))
                            maps.Add((key, a1, MapFieldType(typ)));
                    }
                }
            }

            using var wb = new XLWorkbook(templateExcelPath);
            var ws = wb.Worksheets?.FirstOrDefault()
                     ?? throw new InvalidOperationException("Template has no worksheet");

            foreach (var m in maps)
            {
                if (!inputs.TryGetValue(m.key, out var raw)) continue;
                var cell = ws.Cell(m.a1);

                if (string.IsNullOrEmpty(raw))
                {
                    cell.Clear(XLClearOptions.Contents);
                    continue;
                }

                if (raw.Contains('\n') || raw.Contains('\r')) cell.Style.Alignment.WrapText = true;

                switch (m.type.ToLowerInvariant())
                {
                    case "date":
                        if (DateTime.TryParse(raw, out var dt))
                        {
                            cell.Value = dt;
                            cell.Style.DateFormat.Format = "yyyy-MM-dd";
                        }
                        else cell.Value = raw;
                        break;
                    case "num":
                        if (decimal.TryParse(raw, out var dec)) cell.Value = dec;
                        else cell.Value = raw;
                        break;
                    default:
                        cell.Value = raw;
                        break;
                }
            }

            var outDir = Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "Docs");
            Directory.CreateDirectory(outDir);

            var docId = GenerateDocumentId();
            var outPath = Path.Combine(outDir, $"{docId}.xlsx");

            wb.SaveAs(outPath);
            return outPath;

            static string A1FromRowCol(int row, int col)
            {
                if (row < 1 || col < 1) return "";
                string letters = "";
                int n = col;
                while (n > 0)
                {
                    int r = (n - 1) % 26;
                    letters = (char)('A' + r) + letters;
                    n = (n - 1) / 26;
                }
                return $"{letters}{row}";
            }

            static string MapFieldType(string? t)
            {
                var s = (t ?? "").Trim().ToLowerInvariant();
                if (s.StartsWith("date")) return "date";
                if (s.StartsWith("num") || s.Contains("number") || s.Contains("decimal") || s.Contains("integer")) return "num";
                return "text";
            }
        }

        private static string GenerateDocumentId()
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string suffix = RandomAlphaNumUpper(4);
            return $"DOC_{timestamp}{suffix}";
        }

        private static string RandomAlphaNumUpper(int len)
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            Span<byte> buffer = stackalloc byte[len];
            RandomNumberGenerator.Fill(buffer);
            char[] chars = new char[len];
            for (int i = 0; i < len; i++)
                chars[i] = alphabet[buffer[i] % alphabet.Length];
            return new string(chars);
        }


        // ========= Descriptor 변환/FlowGroups =========
        private static string ConvertDescriptorIfNeeded(string? json)
        {
            const string EMPTY = "{\"inputs\":[],\"approvals\":[],\"version\":\"converted\"}";
            if (!TryParseJsonFlexible(json, out var doc)) return EMPTY;

            try
            {
                var root = doc.RootElement;
                if (root.TryGetProperty("inputs", out _)) return json!;

                List<(string key, string type, string a1)> inputs = new();
                if (root.TryGetProperty("Fields", out var fieldsEl) && fieldsEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var f in fieldsEl.EnumerateArray())
                    {
                        var key = f.TryGetProperty("Key", out var k) ? (k.GetString() ?? "").Trim() : "";
                        if (string.IsNullOrEmpty(key)) continue;
                        var type = f.TryGetProperty("Type", out var t) ? (t.GetString() ?? "Text") : "Text";
                        string a1 = "";
                        if (f.TryGetProperty("Cell", out var cell) && cell.ValueKind == JsonValueKind.Object)
                            a1 = cell.TryGetProperty("A1", out var a) ? (a.GetString() ?? "") : "";
                        inputs.Add((key, MapType(type), a1));
                    }
                }

                // 2025.12.12 Changed: approvalsDict (Slot 기반 Dictionary) 사용을 제거하고
                //                     Approvals 배열 순서대로 A1, A2, ... 를 부여하는 리스트 기반으로 변환
                //                     Slot 값이 모두 1 이더라도 결재자 2명 이상이 그대로 유지되도록 수정
                // 2025.12.12 Removed: var approvalsDict = new Dictionary<int, (string approverType, string? value)>();
                //                     approvalsDict[slot] = (MapApproverType(typ), val);
                //                     approvalsDict.OrderBy(kv => kv.Key) ...
                var approvals = new List<object>();
                if (root.TryGetProperty("Approvals", out var apprEl) && apprEl.ValueKind == JsonValueKind.Array)
                {
                    int index = 0;
                    foreach (var a in apprEl.EnumerateArray())
                    {
                        // Slot 은 더 이상 step 키로 사용하지 않고, 원본 배열 순서(index)를 그대로 사용
                        var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                        var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;

                        var roleKey = $"A{index + 1}"; // A1, A2, A3 ...
                        approvals.Add(new
                        {
                            roleKey,
                            approverType = MapApproverType(typ),
                            required = false,
                            value = val ?? string.Empty
                        });

                        index++;
                    }
                }

                var obj = new
                {
                    inputs = inputs.Select(x => new { key = x.key, type = x.type, required = false, a1 = x.a1 }).ToList(),
                    // 2025.12.12 Changed: 위에서 구성한 approvals 리스트 사용
                    approvals,
                    version = "converted"
                };
                return JsonSerializer.Serialize(obj);
            }
            catch
            {
                return EMPTY;
            }
            finally
            {
                doc.Dispose();
            }

            static string MapType(string t)
            {
                t = (t ?? "").Trim().ToLowerInvariant();
                if (t.StartsWith("date")) return "Date";
                if (t.StartsWith("num") || t.Contains("number") || t.Contains("decimal") || t.Contains("integer")) return "Num";
                return "Text";
            }

            static string MapApproverType(string t)
            {
                t = (t ?? "").Trim();
                return (t == "Person" || t == "Role" || t == "Rule") ? t : "Person";
            }
        }

        private string BuildDescriptorJsonWithFlowGroups(string templateCode, string rawDescriptorJson)
        {
            // 1) 기존 변환 로직 그대로 유지
            var normalized = ConvertDescriptorIfNeeded(rawDescriptorJson);

            var versionId = GetLatestVersionId(templateCode);
            if (versionId <= 0) return normalized;

            var groups = BuildFlowGroupsForTemplate(versionId);

            DescriptorDto desc;
            try
            {
                desc = JsonSerializer.Deserialize<DescriptorDto>(normalized) ?? new DescriptorDto();
            }
            catch
            {
                desc = new DescriptorDto();
            }

            // 2) 템플릿 원본(rawDescriptorJson)에 Fields/Approvals 가 있으면
            //    그 Approvals 를 기준으로 A1, A2, ... 구조를 다시 만들어서 덮어씀
            desc.Approvals = RebuildApprovalsFromLegacyDescriptor(rawDescriptorJson, desc.Approvals);

            desc.FlowGroups = groups;

            var json = JsonSerializer.Serialize(
                desc,
                new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                    WriteIndented = false
                });

            // ★ 추가: approvals 배열에서 Step 기준(A1/order/Slot) 중복 제거
            return DedupApprovalsJsonByStep(json);
        }

        private static string DedupApprovalsJsonByStep(string json)
        {
            try
            {
                var root = System.Text.Json.Nodes.JsonNode.Parse(json) as System.Text.Json.Nodes.JsonObject;
                if (root == null) return json;

                // approvals / Approvals 둘 다 지원
                var approvalsNode = root["approvals"] ?? root["Approvals"];
                if (approvalsNode is not System.Text.Json.Nodes.JsonArray arr) return json;

                static int GetStep(System.Text.Json.Nodes.JsonNode? n)
                {
                    if (n is not System.Text.Json.Nodes.JsonObject o) return 0;

                    if (o.TryGetPropertyValue("Slot", out var slot) && slot != null)
                    {
                        if (slot is System.Text.Json.Nodes.JsonValue v1 && v1.TryGetValue<int>(out var si)) return si;
                        if (slot is System.Text.Json.Nodes.JsonValue v2 && v2.TryGetValue<string>(out var ss) && int.TryParse(ss, out var s2)) return s2;
                    }

                    if (o.TryGetPropertyValue("order", out var ord) && ord != null)
                    {
                        if (ord is System.Text.Json.Nodes.JsonValue v1 && v1.TryGetValue<int>(out var oi)) return oi;
                        if (ord is System.Text.Json.Nodes.JsonValue v2 && v2.TryGetValue<string>(out var os) && int.TryParse(os, out var o2)) return o2;
                    }

                    string? rk = null;
                    if (o.TryGetPropertyValue("roleKey", out var rkn) && rkn is System.Text.Json.Nodes.JsonValue rv && rv.TryGetValue<string>(out var rks))
                        rk = rks;
                    if (string.IsNullOrWhiteSpace(rk) && o.TryGetPropertyValue("RoleKey", out var rkn2) && rkn2 is System.Text.Json.Nodes.JsonValue rv2 && rv2.TryGetValue<string>(out var rks2))
                        rk = rks2;

                    if (!string.IsNullOrWhiteSpace(rk))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(rk.Trim(), @"(?i)^A(\d+)$");
                        if (m.Success && int.TryParse(m.Groups[1].Value, out var rs)) return rs;
                    }

                    return 0;
                }

                static bool HasAnyValue(System.Text.Json.Nodes.JsonNode? n)
                {
                    if (n is not System.Text.Json.Nodes.JsonObject o) return false;

                    static string? GetStr(System.Text.Json.Nodes.JsonObject o2, string key)
                    {
                        if (o2.TryGetPropertyValue(key, out var v) && v is System.Text.Json.Nodes.JsonValue jv && jv.TryGetValue<string>(out var s))
                            return s;
                        return null;
                    }

                    if (!string.IsNullOrWhiteSpace(GetStr(o, "ApproverValue"))) return true;
                    if (!string.IsNullOrWhiteSpace(GetStr(o, "value"))) return true;
                    if (!string.IsNullOrWhiteSpace(GetStr(o, "email"))) return true;

                    if (o.TryGetPropertyValue("emails", out var e) && e is System.Text.Json.Nodes.JsonArray ea && ea.Count > 0) return true;
                    if (o.TryGetPropertyValue("users", out var us) && us is System.Text.Json.Nodes.JsonArray ua && ua.Count > 0) return true;
                    if (o.TryGetPropertyValue("user", out var uo) && uo is System.Text.Json.Nodes.JsonObject) return true;

                    return false;
                }

                // Step별로 “가장 좋은” 항목만 유지(값 있는 항목 우선)
                var bestByStep = new Dictionary<int, System.Text.Json.Nodes.JsonNode?>();
                foreach (var item in arr)
                {
                    var step = GetStep(item);
                    if (step <= 0) continue;

                    if (!bestByStep.TryGetValue(step, out var existing))
                    {
                        bestByStep[step] = item;
                    }
                    else
                    {
                        if (!HasAnyValue(existing) && HasAnyValue(item))
                            bestByStep[step] = item;
                    }
                }

                var rebuilt = new System.Text.Json.Nodes.JsonArray();
                foreach (var kv in bestByStep.OrderBy(k => k.Key))
                    rebuilt.Add(kv.Value);

                if (root.ContainsKey("approvals")) root["approvals"] = rebuilt;
                else if (root.ContainsKey("Approvals")) root["Approvals"] = rebuilt;

                return root.ToJsonString(new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                    WriteIndented = false
                });
            }
            catch
            {
                return json;
            }
        }

        // 2025.12.12 Added: 템플릿의 옛 형식 DescriptorJson(Fields/Approvals)을 읽어
        //                   approvals 리스트를 A1, A2,... 형태로 재구성하는 헬퍼
        private static List<object>? RebuildApprovalsFromLegacyDescriptor(string? legacyJson, List<object>? current)
        {
            // legacyJson 이 없으면 기존 값 유지
            if (!TryParseJsonFlexible(legacyJson, out var doc))
                return current;

            try
            {
                var root = doc.RootElement;

                // 템플릿 쪽 옛 포맷: "Approvals": [ { "Slot":1, "ApproverType":"Person", "ApproverValue":"..." }, ... ]
                if (!root.TryGetProperty("Approvals", out var apprEl) ||
                    apprEl.ValueKind != JsonValueKind.Array)
                    return current;

                var list = new List<object>();
                var index = 1;

                foreach (var a in apprEl.EnumerateArray())
                {
                    var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                    var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;

                    // A1, A2, A3... 으로 강제 부여 (Slot 값이 전부 1이어도 상관없이 순서대로)
                    var roleKey = $"A{index}";

                    // ApproverType 은 Person / Role / Rule 만 허용
                    var mappedType = (typ == "Person" || typ == "Role" || typ == "Rule") ? typ : "Person";

                    list.Add(new
                    {
                        roleKey,
                        approverType = mappedType,
                        required = false,
                        value = val ?? string.Empty
                    });

                    index++;
                }

                // 템플릿에 승인자가 없다면 기존 값 유지
                if (list.Count == 0)
                    return current;

                // 템플릿 정의를 우선시하여 approvals 전체를 교체
                return list;
            }
            catch
            {
                return current;
            }
            finally
            {
                doc.Dispose();
            }
        }

        private long GetLatestVersionId(string templateCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            if (string.IsNullOrWhiteSpace(cs)) return 0;

            try
            {
                using var conn = new SqlConnection(cs);
                conn.Open();
                using var cmd = new SqlCommand(@"
SELECT TOP 1
    COALESCE(CAST(v.Id AS BIGINT), CAST(v.VersionId AS BIGINT)) AS VersionId
FROM DocTemplateMaster m
JOIN DocTemplateVersion v ON v.TemplateId = m.Id
WHERE m.DocCode = @code
ORDER BY v.VersionNo DESC;", conn);
                cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode ?? string.Empty });
                var obj = cmd.ExecuteScalar();
                if (obj is long l) return l;
                if (obj is int i) return i;
                if (obj != null && long.TryParse(obj.ToString(), out var p)) return p;
            }
            catch { }
            return 0;
        }

        private List<FlowGroupDto> BuildFlowGroupsForTemplate(long versionId)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var rows = new List<(string Key, string? A1, int Row, int Col, string BaseKey, int? Index)>();

            using (var conn = new SqlConnection(cs))
            {
                conn.Open();
                using var cmd = new SqlCommand(@"
SELECT [Key], A1, CellRow, CellColumn
FROM dbo.DocTemplateField
WHERE VersionId = @vid AND [Key] IS NOT NULL
ORDER BY [Key], A1;", conn);
                cmd.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = versionId });
                using var rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    var key = rd["Key"] as string ?? "";
                    var a1 = rd["A1"] as string;
                    var r = rd["CellRow"] is DBNull ? 0 : Convert.ToInt32(rd["CellRow"]);
                    var c = rd["CellColumn"] is DBNull ? 0 : Convert.ToInt32(rd["CellColumn"]);
                    var (baseKey, idx) = ParseKey(key);
                    rows.Add((key, a1, r, c, baseKey, idx));
                }
            }

            var groups = rows
                .GroupBy(x => x.BaseKey)
                .Select(g =>
                {
                    var ordered = g
                        .OrderBy(x => x.Index.HasValue ? 0 : 1)
                        .ThenBy(x => x.Index)
                        .ThenBy(x => x.A1)
                        .ThenBy(x => x.Row)
                        .ThenBy(x => x.Col)
                        .Select(x => x.Key)
                        .ToList();

                    return new FlowGroupDto { ID = g.Key, Keys = ordered };
                })
                .Where(g => g.Keys.Count > 1)
                .ToList();

            return groups;
        }

        private static string ExtractApprovalsJson(string descriptorJson)
        {
            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                if (doc.RootElement.TryGetProperty("approvals", out var arr))
                    return arr.GetRawText();
                if (doc.RootElement.TryGetProperty("Approvals", out var arr2))
                    return arr2.GetRawText();
            }
            catch { }
            return "[]";
        }

        private List<string> ResolveApproverEmails(string? descriptorJson, int stepIndex)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(descriptorJson) || stepIndex < 1) return result;

            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                var root = doc.RootElement;

                if (!root.TryGetProperty("approvals", out var approvals))
                {
                    if (root.TryGetProperty("Approvals", out var alt))
                        approvals = alt;
                    else return result;
                }

                if (approvals.ValueKind != JsonValueKind.Array) return result;
                if (approvals.GetArrayLength() < stepIndex) return result;

                var target = approvals[stepIndex - 1];
                var approverType = target.TryGetProperty("approverType", out var at) ? at.GetString() ?? "" : (target.TryGetProperty("ApproverType", out var at2) ? at2.GetString() ?? "" : "");
                var rawValue = target.TryGetProperty("value", out var vv) ? vv.GetString() ?? "" : (target.TryGetProperty("ApproverValue", out var v2) ? v2.GetString() ?? "" : "");

                if (!string.Equals(approverType, "Person", StringComparison.OrdinalIgnoreCase))
                    return result;

                if (LooksLikeEmail(rawValue)) { result.Add(rawValue); return result; }

                try
                {
                    var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                    using var conn = new SqlConnection(cs);
                    conn.Open();

                    using var cmd = new SqlCommand(@"
SELECT TOP 1 Email
FROM dbo.AspNetUsers
WHERE Id = @v OR UserName = @v OR NormalizedUserName = UPPER(@v) OR Email = @v OR NormalizedEmail = UPPER(@v);", conn);
                    cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = rawValue });

                    var emailObj = cmd.ExecuteScalar();
                    var email = emailObj as string;
                    if (!string.IsNullOrWhiteSpace(email) && LooksLikeEmail(email))
                        result.Add(email);
                }
                catch { }
            }
            catch { }

            return result;

            static bool LooksLikeEmail(string s) => !string.IsNullOrWhiteSpace(s) && s.Contains("@") && s.Contains(".");
        }

        private void ApplyStampToExcel(string outXlsx, int slot, string action, string approverDisplay)
        {
            using var wb = new XLWorkbook(outXlsx);
            var ws = wb.Worksheets.First();

            var row1 = 1 + (slot - 1) * 3;
            var col1 = 1;
            var row2 = row1 + 2;
            var col2 = 4;

            var range = ws.Range(row1, col1, row2, col2);
            range.Clear(XLClearOptions.Contents);
            range.Merge();

            var cell = range.FirstCell();
            var now = DateTime.Now;
            var label = action.ToLowerInvariant() switch
            {
                "approve" => $"APPROVED {now:yyyy-MM-dd HH:mm} / {approverDisplay}",
                "hold" => $"ON HOLD {now:yyyy-MM-dd HH:mm} / {approverDisplay}",
                "reject" => $"REJECTED {now:yyyy-MM-dd HH:mm} / {approverDisplay}",
                _ => $"UPDATED {now:yyyy-MM-dd HH:mm} / {approverDisplay}"
            };

            cell.Value = label;

            cell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;

            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 14;

            if (action.Equals("approve", StringComparison.OrdinalIgnoreCase))
            {
                cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGreen;
                cell.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thick;
            }
            else if (action.Equals("hold", StringComparison.OrdinalIgnoreCase))
            {
                cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightYellow;
                cell.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thick;
            }
            else if (action.Equals("reject", StringComparison.OrdinalIgnoreCase))
            {
                cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightPink;
                cell.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thick;
            }

            wb.Save();
        }

        private string ResolveOutputPathFromDocId(string docId)
        {
            var outDir = Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "Docs");
            var path = Path.Combine(outDir, $"{docId}.xlsx");
            return path;
        }
        
        private async Task SaveDocumentAsync(string docId, string templateCode, string title, string status, string outputPath, Dictionary<string, string> inputs, string approvalsJson, string descriptorJson, string userId, string userName, string compCd, string? departmentId, int? templateVersionId = null
)
        {
            if (string.IsNullOrWhiteSpace(compCd))
                throw new InvalidOperationException("CompCd is required by schema but missing.");

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            await using var tx = await conn.BeginTransactionAsync();
            try
            {
                // ---------------- 0) TemplateVersionId 해석 ----------------
                int? effectiveTemplateVersionId = templateVersionId;

                if (!effectiveTemplateVersionId.HasValue)
                {
                    await using var vCmd = conn.CreateCommand();
                    vCmd.Transaction = (SqlTransaction)tx;
                    vCmd.CommandText = @"
SELECT TOP (1) v.Id
FROM   dbo.DocTemplateVersion AS v
JOIN   dbo.DocTemplateMaster  AS m ON v.TemplateId = m.Id
WHERE  m.DocCode = @TemplateCode
ORDER BY v.VersionNo DESC, v.Id DESC;";
                    vCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode });

                    var obj = await vCmd.ExecuteScalarAsync();
                    if (obj != null && obj != DBNull.Value && int.TryParse(obj.ToString(), out var vid))
                        effectiveTemplateVersionId = vid;
                }

                // ---------------- 1) Documents INSERT ----------------
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = (SqlTransaction)tx;
                    cmd.CommandText = @"
INSERT INTO dbo.Documents
    (DocId, TemplateCode, TemplateVersionId, TemplateTitle, Status, OutputPath, CompCd, DepartmentId, CreatedBy, CreatedByName, CreatedAt, DescriptorJson)
VALUES
    (@DocId, @TemplateCode, @TemplateVersionId, @TemplateTitle, @Status, @OutputPath, @CompCd, @DepartmentId, @CreatedBy, @CreatedByName, SYSUTCDATETIME(), @DescriptorJson);";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    cmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode });
                    cmd.Parameters.Add(new SqlParameter("@TemplateVersionId", SqlDbType.Int) { Value = (object?)effectiveTemplateVersionId ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@TemplateTitle", SqlDbType.NVarChar, 400) { Value = (object?)title ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@Status", SqlDbType.NVarChar, 20) { Value = status });
                    cmd.Parameters.Add(new SqlParameter("@OutputPath", SqlDbType.NVarChar, 500) { Value = outputPath });
                    cmd.Parameters.Add(new SqlParameter("@CompCd", SqlDbType.VarChar, 10) { Value = compCd });
                    cmd.Parameters.Add(new SqlParameter("@DepartmentId", SqlDbType.VarChar, 12) { Value = string.IsNullOrWhiteSpace(departmentId) ? DBNull.Value : departmentId });
                    cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 450) { Value = userId ?? "" });
                    cmd.Parameters.Add(new SqlParameter("@CreatedByName", SqlDbType.NVarChar, 200) { Value = userName ?? "" });
                    cmd.Parameters.Add(new SqlParameter("@DescriptorJson", SqlDbType.NVarChar, -1) { Value = (object?)descriptorJson ?? DBNull.Value });
                    await cmd.ExecuteNonQueryAsync();
                }

                // ---------------- 2) DocumentInputs INSERT ----------------
                if (inputs?.Count > 0)
                {
                    foreach (var kv in inputs)
                    {
                        if (string.IsNullOrWhiteSpace(kv.Key)) continue;

                        await using var icmd = conn.CreateCommand();
                        icmd.Transaction = (SqlTransaction)tx;
                        icmd.CommandText = @"
INSERT INTO dbo.DocumentInputs (DocId, FieldKey, FieldValue, CreatedAt)
VALUES (@DocId, @FieldKey, @FieldValue, SYSUTCDATETIME());";
                        icmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        icmd.Parameters.Add(new SqlParameter("@FieldKey", SqlDbType.NVarChar, 200) { Value = kv.Key });
                        icmd.Parameters.Add(new SqlParameter("@FieldValue", SqlDbType.NVarChar, -1) { Value = (object?)kv.Value ?? DBNull.Value });
                        await icmd.ExecuteNonQueryAsync();
                    }
                }

                // ---------------- 3) DocumentApprovals INSERT (approvalsJson 기반) ----------------
                // ★ 변경: value/ApproverValue 뿐 아니라 users/user/emails/email 까지 후보를 추출
                // ★ 변경: step++ 연속 증가가 아니라, 파싱된 StepOrder(step)를 그대로 사용
                try
                {
                    static string Norm(string? s) => (s ?? string.Empty).Trim();

                    static IEnumerable<string> ExtractCandidateValues(JsonElement a)
                    {
                        string? val = null;
                        if (a.TryGetProperty("ApproverValue", out var pAV) && pAV.ValueKind == JsonValueKind.String) val = pAV.GetString();
                        if (string.IsNullOrWhiteSpace(val) && a.TryGetProperty("value", out var pV) && pV.ValueKind == JsonValueKind.String) val = pV.GetString();
                        if (string.IsNullOrWhiteSpace(val) && a.TryGetProperty("Value", out var pV2) && pV2.ValueKind == JsonValueKind.String) val = pV2.GetString();
                        if (!string.IsNullOrWhiteSpace(val)) yield return val!.Trim();

                        if (a.TryGetProperty("email", out var em) && em.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(em.GetString()))
                            yield return em.GetString()!.Trim();

                        if (a.TryGetProperty("emails", out var emails) && emails.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var e in emails.EnumerateArray())
                                if (e.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(e.GetString()))
                                    yield return e.GetString()!.Trim();
                        }

                        if (a.TryGetProperty("user", out var user) && user.ValueKind == JsonValueKind.Object)
                        {
                            if (user.TryGetProperty("email", out var ue) && ue.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(ue.GetString()))
                                yield return ue.GetString()!.Trim();
                            else if (user.TryGetProperty("mail", out var um) && um.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(um.GetString()))
                                yield return um.GetString()!.Trim();
                            else if (user.TryGetProperty("Email", out var uE) && uE.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(uE.GetString()))
                                yield return uE.GetString()!.Trim();
                            else if (user.TryGetProperty("id", out var uid) && uid.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(uid.GetString()))
                                yield return uid.GetString()!.Trim();
                        }

                        if (a.TryGetProperty("users", out var users) && users.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var u in users.EnumerateArray())
                            {
                                if (u.ValueKind != JsonValueKind.Object) continue;

                                if (u.TryGetProperty("email", out var ue) && ue.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(ue.GetString()))
                                    yield return ue.GetString()!.Trim();
                                else if (u.TryGetProperty("mail", out var um) && um.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(um.GetString()))
                                    yield return um.GetString()!.Trim();
                                else if (u.TryGetProperty("Email", out var uE) && uE.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(uE.GetString()))
                                    yield return uE.GetString()!.Trim();
                                else if (u.TryGetProperty("id", out var uid) && uid.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(uid.GetString()))
                                    yield return uid.GetString()!.Trim();
                            }
                        }
                    }

                    static int GetStep(JsonElement a, string roleKey)
                    {
                        int step = 0;

                        if (a.TryGetProperty("Slot", out var pSlot))
                        {
                            if (pSlot.ValueKind == JsonValueKind.Number) pSlot.TryGetInt32(out step);
                            else if (pSlot.ValueKind == JsonValueKind.String) int.TryParse(pSlot.GetString(), out step);
                        }

                        if (step <= 0 && a.TryGetProperty("order", out var pOrd))
                        {
                            if (pOrd.ValueKind == JsonValueKind.Number) pOrd.TryGetInt32(out step);
                            else if (pOrd.ValueKind == JsonValueKind.String) int.TryParse(pOrd.GetString(), out step);
                        }

                        if (step <= 0 && !string.IsNullOrWhiteSpace(roleKey))
                        {
                            var m = System.Text.RegularExpressions.Regex.Match(roleKey.Trim(), @"(?i)^A(\d+)$");
                            if (m.Success) int.TryParse(m.Groups[1].Value, out step);
                        }

                        return step;
                    }

                    using var aj = JsonDocument.Parse(approvalsJson ?? "[]");
                    var arr = aj.RootElement;

                    var toInsert = new List<(int step, string roleKey, string approverValue)>();
                    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var el in arr.EnumerateArray())
                    {
                        if (el.ValueKind != JsonValueKind.Object) continue;

                        string roleKey = "";
                        if (el.TryGetProperty("roleKey", out var rk) && rk.ValueKind == JsonValueKind.String) roleKey = rk.GetString() ?? "";
                        else if (el.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? "";

                        var step = GetStep(el, roleKey);
                        if (step <= 0) step = 1;

                        var rkNorm = string.IsNullOrWhiteSpace(roleKey) ? $"A{step}" : roleKey.Trim();

                        // 후보 값 중 첫 번째 “유효값”만 채택(현재 요구사항: 1 step = 1 승인자)
                        var approverVal = ExtractCandidateValues(el)
                            .Select(Norm)
                            .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

                        if (string.IsNullOrWhiteSpace(approverVal))
                            continue;

                        var key = $"{step}|{rkNorm}|{approverVal}";
                        if (!seen.Add(key)) continue;

                        toInsert.Add((step, rkNorm, approverVal));
                    }

                    foreach (var item in toInsert.OrderBy(x => x.step))
                    {
                        string? resolvedUserId = await TryResolveUserIdAsync(conn, (SqlTransaction)tx, item.approverValue);

                        await using var acmd = conn.CreateCommand();
                        acmd.Transaction = (SqlTransaction)tx;
                        acmd.CommandText = @"
INSERT INTO dbo.DocumentApprovals (DocId, StepOrder, RoleKey, ApproverValue, UserId, Status, CreatedAt)
VALUES (@DocId, @StepOrder, @RoleKey, @ApproverValue, @UserId, 'Pending', SYSUTCDATETIME());";
                        acmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        acmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = item.step });
                        acmd.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 10) { Value = item.roleKey });
                        acmd.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 256) { Value = item.approverValue });
                        acmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)resolvedUserId ?? DBNull.Value });

                        await acmd.ExecuteNonQueryAsync();
                    }
                }
                catch
                {
                    // approvalsJson 파싱 실패해도 문서 저장 자체는 막지 않음
                }

                await tx.CommitAsync();
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }
        }

        private static async Task<int?> ResolveTemplateVersionIdAsync(SqlConnection conn, SqlTransaction tx, string templateCode)
        {
            if (string.IsNullOrWhiteSpace(templateCode))
                return null;

            await using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = @"
SELECT TOP (1) v.Id
FROM dbo.DocTemplateMaster  AS m
JOIN dbo.DocTemplateVersion AS v
     ON v.TemplateId = m.Id
WHERE m.DocCode = @DocCode
ORDER BY v.VersionNo DESC, v.Id DESC;";

            cmd.Parameters.Add(new SqlParameter("@DocCode", SqlDbType.NVarChar, 100)
            {
                Value = templateCode
            });

            var obj = await cmd.ExecuteScalarAsync();
            if (obj == null || obj == DBNull.Value)
                return null;

            return Convert.ToInt32(obj);
        }

        private async Task EnsureApprovalsAndSyncAsync(string docId, string docCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();
            await using var tx = await conn.BeginTransactionAsync();

            try
            {
                // 1) 해당 문서의 DescriptorJson 스냅샷 가져오기
                string? descriptorJson = null;

                await using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = (SqlTransaction)tx;
                    cmd.CommandText = @"
SELECT TOP (1) DescriptorJson
FROM dbo.Documents
WHERE DocId = @DocId;";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                    var obj = await cmd.ExecuteScalarAsync();
                    if (obj != null && obj != DBNull.Value)
                        descriptorJson = Convert.ToString(obj);
                }

                // DescriptorJson 이 없으면 더 할 수 있는 게 없음
                if (string.IsNullOrWhiteSpace(descriptorJson))
                {
                    await tx.CommitAsync();
                    return;
                }

                // 2) descriptorJson → approvals 배열 파싱
                //    - value가 빈 항목은 생성 대상에서 제외(결재자 없는 Pending 라인 방지)
                var approvals = new List<(int StepOrder, string RoleKey, string? ApproverValue)>();

                try
                {
                    using var dj = JsonDocument.Parse(descriptorJson);
                    var root = dj.RootElement;

                    if (root.TryGetProperty("approvals", out var arr) &&
                        arr.ValueKind == JsonValueKind.Array)
                    {
                        var i = 0;
                        foreach (var a in arr.EnumerateArray())
                        {
                            // ApproverValue : value / approverValue / ApproverValue 중 첫 번째 문자열
                            string? approverValue = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String)
                                approverValue = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String)
                                approverValue = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String)
                                approverValue = v3.GetString();

                            // value가 비어있으면 라인 생성 대상에서 제외
                            if (string.IsNullOrWhiteSpace(approverValue))
                            {
                                i++;
                                continue;
                            }

                            // StepOrder : order / slot / index+1 중 우선순위로 결정
                            int stepOrder = i + 1;

                            if (a.TryGetProperty("order", out var ordEl))
                            {
                                if (ordEl.ValueKind == JsonValueKind.Number &&
                                    ordEl.TryGetInt32(out var ordNum))
                                    stepOrder = ordNum;
                                else if (ordEl.ValueKind == JsonValueKind.String &&
                                         int.TryParse(ordEl.GetString(), out var ordStr))
                                    stepOrder = ordStr;
                            }
                            else if (a.TryGetProperty("slot", out var slotEl))
                            {
                                if (slotEl.ValueKind == JsonValueKind.Number &&
                                    slotEl.TryGetInt32(out var slotNum))
                                    stepOrder = slotNum;
                                else if (slotEl.ValueKind == JsonValueKind.String &&
                                         int.TryParse(slotEl.GetString(), out var slotStr))
                                    stepOrder = slotStr;
                            }

                            // RoleKey : roleKey / RoleKey / part 우선
                            string roleKey = string.Empty;

                            if (a.TryGetProperty("roleKey", out var rk1) && rk1.ValueKind == JsonValueKind.String)
                                roleKey = rk1.GetString() ?? string.Empty;
                            else if (a.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String)
                                roleKey = rk2.GetString() ?? string.Empty;
                            else if (a.TryGetProperty("part", out var rk3) && rk3.ValueKind == JsonValueKind.String)
                                roleKey = rk3.GetString() ?? string.Empty;

                            if (string.IsNullOrWhiteSpace(roleKey))
                                roleKey = $"A{stepOrder}";

                            approvals.Add((stepOrder, roleKey, approverValue!.Trim()));
                            i++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "EnsureApprovalsAndSyncAsync: descriptorJson 파싱 실패 docId={docId}", docId);
                }

                // approvals 정의가 전혀 없으면 승인행 생성 불필요
                if (approvals.Count == 0)
                {
                    await tx.CommitAsync();
                    return;
                }

                // 3) 기존 DocumentApprovals 모두 삭제
                await using (var del = conn.CreateCommand())
                {
                    del.Transaction = (SqlTransaction)tx;
                    del.CommandText = @"DELETE FROM dbo.DocumentApprovals WHERE DocId = @DocId;";
                    del.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    await del.ExecuteNonQueryAsync();
                }

                // 4) descriptor 기준으로 DocumentApprovals 재생성
                var ordered = approvals
                    .OrderBy(a => a.StepOrder)
                    .ThenBy(a => a.RoleKey, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // StepOrder가 중복/비정상이어도 최종 저장은 1..N으로 정규화(안전)
                for (int idx = 0; idx < ordered.Count; idx++)
                {
                    var a = ordered[idx];
                    var stepOrder = idx + 1;               // 정규화
                    var roleKey = $"A{stepOrder}";         // 정규화

                    await using var ins = conn.CreateCommand();
                    ins.Transaction = (SqlTransaction)tx;
                    ins.CommandText = @"
INSERT INTO dbo.DocumentApprovals
    (DocId, StepOrder, RoleKey, ApproverValue, Status, CreatedAt)
VALUES
    (@DocId, @StepOrder, @RoleKey, @ApproverValue, N'Pending', SYSUTCDATETIME());";

                    ins.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    ins.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = stepOrder });
                    ins.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 100) { Value = roleKey });

                    var pVal = new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128);
                    pVal.Value = a.ApproverValue!;
                    ins.Parameters.Add(pVal);

                    await ins.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();

                _log.LogInformation(
                    "EnsureApprovalsAndSyncAsync completed docId={docId}, docCode={docCode}, count={cnt}",
                    docId, docCode, ordered.Count);
            }
            catch (Exception ex)
            {
                try { await tx.RollbackAsync(); } catch { }
                _log.LogError(ex, "EnsureApprovalsAndSyncAsync failed docId={docId}, docCode={docCode}", docId, docCode);
            }
        }

        private static string EnsureDescriptorHasApprovals(string descriptorJson, string approvalsJson)
        {
            if (string.IsNullOrWhiteSpace(descriptorJson))
                return descriptorJson;

            if (string.IsNullOrWhiteSpace(approvalsJson) || approvalsJson.Trim() == "[]")
                return descriptorJson;

            try
            {
                using var dj = System.Text.Json.JsonDocument.Parse(descriptorJson);
                var root = dj.RootElement;

                if (root.ValueKind != System.Text.Json.JsonValueKind.Object)
                    return descriptorJson;

                if (root.TryGetProperty("approvals", out var arr) && arr.ValueKind == System.Text.Json.JsonValueKind.Array)
                    return descriptorJson;

                var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

                foreach (var prop in root.EnumerateObject())
                {
                    dict[prop.Name] = System.Text.Json.JsonSerializer.Deserialize<object>(prop.Value.GetRawText());
                }

                dict["approvals"] = System.Text.Json.JsonSerializer.Deserialize<object>(approvalsJson);

                return System.Text.Json.JsonSerializer.Serialize(dict);
            }
            catch
            {
                return descriptorJson;
            }
        }

        private string BuildBoardStatusLabel(string rawStatus, int completedSteps, int totalSteps)
        {
            if (string.IsNullOrWhiteSpace(rawStatus))
                return string.Empty;

            var code = rawStatus.Trim();

            // 2025.12.08 Changed: 이미 DOC_Status_* 형식인 경우 그대로 진행률만 덧붙이도록 유지
            if (code.StartsWith("DOC_Status_", StringComparison.OrdinalIgnoreCase))
            {
                var baseLabel = _S[code].Value;
                return AppendProgress(baseLabel, completedSteps, totalSteps);
            }

            // 2) PendingA1, OnHoldA2, RejectedA1 같은 코드에서 꼬리 "A{n}" 제거 → 기본 코드만 추출
            string baseCode = code;
            var idxA = code.LastIndexOf('A');
            if (idxA > 0 && idxA < code.Length - 1)
            {
                var tail = code.Substring(idxA + 1);
                if (int.TryParse(tail, out _))
                {
                    baseCode = code.Substring(0, idxA); // 예: "PendingA1" → "Pending"
                }
            }

            // 3) 기본 코드 → 리소스 키 매핑
            string resKey = baseCode switch
            {
                "Draft" => "DOC_Status_Draft",
                "Submitted" => "DOC_Status_Submitted",
                "Pending" => "DOC_Status_Pending",
                "OnHold" => "DOC_Status_OnHold",
                "Approved" => "DOC_Status_Approved",
                "Rejected" => "DOC_Status_Rejected",
                "Recalled" => "DOC_Status_Recalled",
                _ => string.Empty
            };

            var label = string.IsNullOrEmpty(resKey)
                ? code
                : _S[resKey].Value;

            return AppendProgress(label, completedSteps, totalSteps);

            static string AppendProgress(string label, int completed, int total)
            {
                if (total <= 0) return label;

                var done = Math.Max(0, Math.Min(completed, total));
                return $"{label} ({done}/{total})";
            }
        }

        private (List<string> emails, List<string> diag) GetInitialRecipients(ComposePostDto dto, string normalizedDesc, IConfiguration cfg, string? docId = null, string? templateCode = null)
        {
            var emails = new List<string>();
            var diag = new List<string>();

            static bool LooksLikeEmail(string s) => !string.IsNullOrWhiteSpace(s) && s.Contains("@") && s.Contains(".");

            void AppendTokens(string raw, SqlConnection? conn)
            {
                if (string.IsNullOrWhiteSpace(raw)) return;
                var tokens = raw.Split(new[] { ';', ',', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                .Distinct(StringComparer.OrdinalIgnoreCase);
                foreach (var tk in tokens)
                {
                    if (LooksLikeEmail(tk)) { emails.Add(tk); diag.Add($"'{tk}' -> email"); continue; }
                    if (conn == null) { diag.Add($"'{tk}' -> no-conn"); continue; }

                    string? email = null;
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
SELECT TOP 1 Email
FROM dbo.AspNetUsers
WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);";
                        cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = tk });
                        email = cmd.ExecuteScalar() as string;
                    }
                    if (!string.IsNullOrWhiteSpace(email) && LooksLikeEmail(email))
                    { emails.Add(email!); diag.Add($"'{tk}' -> AspNetUsers '{email}'"); continue; }

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
SELECT TOP 1 COALESCE(u.Email, p.Email)
FROM dbo.UserProfiles p
LEFT JOIN dbo.AspNetUsers u ON u.Id = p.UserId
WHERE p.DisplayName=@n OR p.Name=@n OR p.UserId=@n OR p.Email=@n;";
                        cmd.Parameters.Add(new SqlParameter("@n", SqlDbType.NVarChar, 256) { Value = tk });
                        email = cmd.ExecuteScalar() as string;
                    }
                    if (!string.IsNullOrWhiteSpace(email) && LooksLikeEmail(email))
                    { emails.Add(email!); diag.Add($"'{tk}' -> UserProfiles '{email}'"); }
                    else
                    { diag.Add($"'{tk}' not resolved"); }
                }
            }

            var cs = cfg.GetConnectionString("DefaultConnection") ?? "";
            using var conn = string.IsNullOrWhiteSpace(cs) ? null : new SqlConnection(cs);
            if (conn != null) conn.Open();

            if (dto?.Mail?.TO is { Count: > 0 }) { foreach (var s in dto.Mail.TO) AppendTokens(s, conn); }
            if (emails.Count == 0 && dto?.Approvals?.To is { Count: > 0 }) { foreach (var s in dto.Approvals.To) AppendTokens(s, conn); }
            if (emails.Count == 0 && dto?.Approvals?.Steps is { Count: > 0 })
                foreach (var st in dto.Approvals.Steps) if (!string.IsNullOrWhiteSpace(st?.Value)) AppendTokens(st!.Value!, conn);

            if (emails.Count == 0)
            {
                try
                {
                    using var doc = JsonDocument.Parse(normalizedDesc);
                    if (doc.RootElement.TryGetProperty("approvals", out var arr) && arr.ValueKind == JsonValueKind.Array)
                        foreach (var el in arr.EnumerateArray())
                        {
                            var val = el.TryGetProperty("value", out var v) ? v.GetString()
                                     : el.TryGetProperty("ApproverValue", out var v2) ? v2.GetString()
                                     : null;
                            if (!string.IsNullOrWhiteSpace(val)) AppendTokens(val!, conn);
                        }
                }
                catch { }
            }

            if (emails.Count == 0 && !string.IsNullOrWhiteSpace(docId) && conn != null)
            {
                try
                {
                    using var q = conn.CreateCommand();
                    q.CommandText = @"SELECT ApproverValue FROM dbo.DocumentApprovals
                              WHERE DocId=@id AND ApproverValue IS NOT NULL AND LTRIM(RTRIM(ApproverValue))<>'';";
                    q.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                    using var rd = q.ExecuteReader();
                    while (rd.Read()) AppendTokens(rd.GetString(0), conn);
                    diag.Add("fallback: DocumentApprovals used");
                }
                catch (Exception ex) { diag.Add("db fallback ex: " + ex.Message); }
            }

            if (emails.Count == 0 && !string.IsNullOrWhiteSpace(templateCode) && conn != null)
            {
                try
                {
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = @"
SELECT TOP (1) v.DescriptorJson
FROM dbo.DocTemplateVersion v
JOIN dbo.DocTemplateMaster  m ON m.Id=v.TemplateId
WHERE m.DocCode=@code
ORDER BY v.VersionNo DESC;";
                    cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode });
                    var raw = cmd.ExecuteScalar() as string;
                    if (!string.IsNullOrWhiteSpace(raw))
                    {
                        using var jd = JsonDocument.Parse(raw);
                        var root = jd.RootElement;
                        var path = root.TryGetProperty("approvals", out var _)
                                   ? "approvals"
                                   : (root.TryGetProperty("Approvals", out var __) ? "Approvals" : null);
                        if (path != null)
                        {
                            var arr = root.GetProperty(path);
                            foreach (var el in arr.EnumerateArray())
                            {
                                var val = el.TryGetProperty("value", out var v) ? v.GetString()
                                         : el.TryGetProperty("ApproverValue", out var v2) ? v2.GetString()
                                         : null;
                                if (!string.IsNullOrWhiteSpace(val)) AppendTokens(val!, conn);
                            }
                            diag.Add("fallback: Template.DescriptorJson approvals used");
                        }
                    }
                }
                catch (Exception ex) { diag.Add("template fallback ex: " + ex.Message); }
            }

            return (emails.Distinct(StringComparer.OrdinalIgnoreCase).ToList(), diag);
        }


        private (string subject, string bodyHtml) BuildSubmissionMail(string authorName, string docTitle, string docId, string? recipientDisplay)
        {
            // 기존 계산 코드는 모두 생략하고, 공통 오버로드로 위임
            return BuildSubmissionMail(authorName, docTitle, docId, recipientDisplay, null);
        }

        private (string subject, string bodyHtml) BuildSubmissionMail(string authorName,string docTitle, string docId, string? recipientDisplay, DateTime? createdAtLocalForMail = null)
        {
            var link = Url.Action("Detail", "Doc", new { id = docId }, Request.Scheme) ?? "";

            var tz = ResolveCompanyTimeZone();
            // createdAtLocalForMail 이 넘어오면 그대로 사용, 없으면 현재 시각을 회사 시간대로 변환
            var nowLocal = createdAtLocalForMail ?? TimeZoneInfo.ConvertTime(DateTime.UtcNow, tz);

            var mm = nowLocal.Month;
            var dd = nowLocal.Day;
            var hh = nowLocal.Hour;
            var min = nowLocal.Minute;

            var subject = string.Format(
                _S["DOC_Email_Subject"].Value,
                authorName,
                docTitle
            );

            var body = string.Format(
                _S["DOC_Email_BodyHtml"].Value,
                string.IsNullOrWhiteSpace(recipientDisplay) ? "" : recipientDisplay, // 인사말 수신자
                authorName,                                                          // 문서 작성자
                docTitle,                                                            // 문서명
                mm, dd, hh, min,                                                     // 작성 시각
                link                                                                 // 상세 링크
            );

            return (subject, body);
        }
        
        private async Task FillDocumentApprovalsFromEmailsAsync(string docId, IEnumerable<string>? approverEmails)
        {
            // 1) 메일 목록 정리
            var list = approverEmails?
                .Where(e => !string.IsNullOrWhiteSpace(e))
                .Select(e => e.Trim())
                .ToList() ?? new List<string>();

            _log.LogInformation("FillDocumentApprovalsFromEmailsAsync START docId={docId}, rawEmails=[{emails}]",
                docId, string.Join(", ", list));

            if (list.Count == 0)
            {
                _log.LogWarning("FillDocumentApprovalsFromEmailsAsync docId={docId} 메일 목록이 비어 있습니다. DocumentApprovals 매핑 건너뜀.", docId);
                return;
            }

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // 2) 현재 DocumentApprovals 상태 로그
            await using (var dumpBefore = conn.CreateCommand())
            {
                dumpBefore.CommandText = @"
SELECT StepOrder, RoleKey, ApproverValue, UserId, Status
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
ORDER BY StepOrder;";
                dumpBefore.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                await using var r = await dumpBefore.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    _log.LogInformation(
                        "DocApprovals BEFORE docId={docId} step={step} role={role} value={val} userId={uid} status={st}",
                        docId,
                        r.GetInt32(0),                         // StepOrder
                        r.GetString(1),                        // RoleKey
                        r.IsDBNull(2) ? "(null)" : r.GetString(2), // ApproverValue
                        r.IsDBNull(3) ? "(null)" : r.GetString(3), // UserId
                        r.IsDBNull(4) ? "(null)" : r.GetString(4)  // Status
                    );
                }
            }

            // 3) StepOrder 1,2,3,… 순서대로 이메일 매핑
            for (var i = 0; i < list.Count; i++)
            {
                var step = i + 1;           // StepOrder = 1부터 시작한다고 가정
                var email = list[i];

                // UserId 조회
                string? userId = null;
                await using (var userCmd = conn.CreateCommand())
                {
                    userCmd.CommandText = @"
SELECT TOP (1) Id
FROM dbo.AspNetUsers
WHERE NormalizedEmail = UPPER(@E)
   OR NormalizedUserName = UPPER(@E)
   OR Id = @E;";
                    userCmd.Parameters.Add(new SqlParameter("@E", SqlDbType.NVarChar, 256) { Value = email });
                    var obj = await userCmd.ExecuteScalarAsync();
                    userId = obj as string;
                }

                _log.LogInformation(
                    "FillDocumentApprovalsFromEmailsAsync RESOLVE docId={docId} step={step} email={email} resolvedUserId={uid}",
                    docId, step, email, userId ?? "(null)");

                // 실제 UPDATE
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
UPDATE a
SET a.ApproverValue = @Email,
    a.UserId        = COALESCE(a.UserId, @UserId)
FROM dbo.DocumentApprovals AS a
WHERE a.DocId     = @DocId
  AND a.StepOrder = @Step
  AND (ISNULL(a.ApproverValue, N'') = N'' OR ISNULL(a.UserId, N'') = N'');";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    cmd.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = step });
                    cmd.Parameters.Add(new SqlParameter("@Email", SqlDbType.NVarChar, 256) { Value = email });
                    cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64)
                    {
                        Value = (object?)userId ?? DBNull.Value
                    });

                    var affected = await cmd.ExecuteNonQueryAsync();
                    _log.LogInformation(
                        "FillDocumentApprovalsFromEmailsAsync UPDATE docId={docId} step={step} email={email} affectedRows={rows}",
                        docId, step, email, affected);
                }
            }

            // 4) 업데이트 이후 상태 로그
            await using (var dumpAfter = conn.CreateCommand())
            {
                dumpAfter.CommandText = @"
SELECT StepOrder, RoleKey, ApproverValue, UserId, Status
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
ORDER BY StepOrder;";
                dumpAfter.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                await using var r2 = await dumpAfter.ExecuteReaderAsync();
                while (await r2.ReadAsync())
                {
                    _log.LogInformation(
                        "DocApprovals AFTER docId={docId} step={step} role={role} value={val} userId={uid} status={st}",
                        docId,
                        r2.GetInt32(0),
                        r2.GetString(1),
                        r2.IsDBNull(2) ? "(null)" : r2.GetString(2),
                        r2.IsDBNull(3) ? "(null)" : r2.GetString(3),
                        r2.IsDBNull(4) ? "(null)" : r2.GetString(4)
                    );
                }
            }

            _log.LogInformation("FillDocumentApprovalsFromEmailsAsync END docId={docId}", docId);
        }

        // ========= UNIQUE 충돌 가드 =========
        private async Task<(string docId, string outputPath)> SaveDocumentWithCollisionGuardAsync(string docId, string templateCode, string title, string status, string outputPath, Dictionary<string, string> inputs, string approvalsJson, string descriptorJson, string userId, string userName, string compCd, string? departmentId)        
        {
            static string WithNewSuffix(string baseDocId)
            {
                var core = string.IsNullOrWhiteSpace(baseDocId)
                    ? "DOC_" + DateTime.Now.ToString("yyyyMMddHHmmssfff")
                    : baseDocId;
                var head = core.Length >= 4 ? core[..^4] : core;
                return head + RandomAlphaNumUpper(4);
            }

            static string EnsurePathFor(string path, string newDocId)
            {
                var dir = string.IsNullOrWhiteSpace(path)
                    ? AppContext.BaseDirectory
                    : (Path.GetDirectoryName(path) ?? AppContext.BaseDirectory);
                var ext = Path.GetExtension(path);
                if (string.IsNullOrWhiteSpace(ext)) ext = ".xlsx";
                return Path.Combine(dir, $"{newDocId}{ext}");
            }

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    await SaveDocumentAsync(
                        docId: docId,
                        templateCode: templateCode,
                        title: title,
                        status: status,
                        outputPath: outputPath,
                        inputs: inputs,
                        approvalsJson: approvalsJson,
                        descriptorJson: descriptorJson,
                        userId: userId,
                        userName: userName,
                        compCd: compCd,
                        departmentId: departmentId
                    );

                    return (docId, outputPath);
                }
                catch (SqlException se) when (se.Number == 2627 || se.Number == 2601)
                {
                    _log.LogWarning(se, "DocId unique violation attempt={attempt} docId={docId}", attempt + 1, docId);

                    var newDocId = WithNewSuffix(docId);
                    var newPath = EnsurePathFor(outputPath, newDocId);

                    if (System.IO.File.Exists(newPath))
                    {
                        newDocId = WithNewSuffix(newDocId);
                        newPath = EnsurePathFor(outputPath, newDocId);
                    }

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(outputPath) && System.IO.File.Exists(outputPath)
                            && !string.Equals(outputPath, newPath, StringComparison.OrdinalIgnoreCase))
                        {
                            System.IO.File.Move(outputPath, newPath, overwrite: false);
                        }
                    }
                    catch (Exception exMove)
                    {
                        _log.LogWarning(exMove, "Document file rename failed old={old} new={@new}", outputPath, newPath);
                    }

                    docId = newDocId;
                    outputPath = newPath;
                    continue;
                }
            }

            throw new InvalidOperationException("DocId collision could not be resolved after retries");
        }
        public sealed class ComposePostDto
        {
            [JsonPropertyName("templateCode")]
            public string? TemplateCode { get; set; }

            [JsonPropertyName("inputs")]
            public Dictionary<string, string>? Inputs { get; set; }

            [JsonPropertyName("approvals")]
            public ApprovalsDto? Approvals { get; set; }

            [JsonPropertyName("mail")]
            public MailDto? Mail { get; set; }

            [JsonPropertyName("descriptorVersion")]
            public string? DescriptorVersion { get; set; }

            // 2025.12.23 Added: Compose 화면 조직 멀티 콤보박스 선택 사용자(공유 대상)
            [JsonPropertyName("selectedRecipientUserIds")]
            public List<string>? SelectedRecipientUserIds { get; set; }

            // 2025.12.23 Added: Compose 첨부파일(임시 업로드 결과) 목록
            [JsonPropertyName("attachments")]
            public List<ComposeAttachmentDto>? Attachments { get; set; }
        }
        public class ComposeAttachmentDto
        {
            [JsonPropertyName("fileKey")]
            public string? FileKey { get; set; }        // /DocFile/Upload 응답의 fileKey

            [JsonPropertyName("originalName")]
            public string? OriginalName { get; set; }   // 원본 파일명(표시용)

            [JsonPropertyName("contentType")]
            public string? ContentType { get; set; }

            [JsonPropertyName("byteSize")]
            public long? ByteSize { get; set; }
        }
        public sealed class ApprovalsDto
        {
            public List<string>? To { get; set; }
            public List<ApprovalStepDto>? Steps { get; set; }
        }


        public sealed class ApprovalStepDto
        {
            public string? RoleKey { get; set; }
            public string? ApproverType { get; set; }
            public string? Value { get; set; }
        }

        public sealed class MailDto
        {
            public List<string>? TO { get; set; }
            public List<string>? CC { get; set; }
            public List<string>? BCC { get; set; }
            public string? Subject { get; set; }
            public string? Body { get; set; }
            public bool Send { get; set; } = true;
            public string? From { get; set; }
        }

        private sealed class DescriptorDto
        {
            public string? Version { get; set; }
            public List<InputFieldDto>? Inputs { get; set; }
            public Dictionary<string, object>? Styles { get; set; }
            public List<object>? Approvals { get; set; }
            public List<FlowGroupDto>? FlowGroups { get; set; }
        }

        private sealed class InputFieldDto
        {
            public string Key { get; set; } = "";
            public string? A1 { get; set; }
            public string? Type { get; set; }
        }

        private sealed class FlowGroupDto
        {
            public string ID { get; set; } = "";
            public List<string> Keys { get; set; } = new();
        }
    }
}
