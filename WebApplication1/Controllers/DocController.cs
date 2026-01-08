using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Globalization;
using System.Threading.Tasks;
using ClosedXML.Excel;
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
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services;
using WebApplication1.Controllers;
using System.Text;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class DocController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IConfiguration _cfg;
        private readonly IAntiforgery _antiforgery;
        private readonly IWebHostEnvironment _env;
        private readonly ApplicationDbContext _db;
        private readonly IDocTemplateService _tpl;
        private readonly ILogger<DocController> _log;
        private readonly IEmailSender _emailSender;
        private readonly SmtpOptions _smtpOpt;
        private static readonly object _attachSeqLock = new object();


        public DocController(
            IStringLocalizer<SharedResource> S,
            IConfiguration cfg,
            IAntiforgery antiforgery,
            IWebHostEnvironment env,
            ApplicationDbContext db,
            IDocTemplateService tpl,
            IEmailSender emailSender,
            IOptions<SmtpOptions> smtpOptions,
            ILogger<DocController> log
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
        }

        private async Task<List<OrgTreeNode>> BuildOrgTreeNodesAsync(string langCode)
        {
            var orgNodes = new List<OrgTreeNode>();
            var compMap = new Dictionary<string, OrgTreeNode>(StringComparer.OrdinalIgnoreCase);
            var deptMap = new Dictionary<string, OrgTreeNode>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT
    a.CompCd,
    ISNULL(g.Name, a.CompCd)        AS CompName,
    a.DepartmentId,
    ISNULL(f.Name, CAST(a.DepartmentId AS nvarchar(50))) AS DeptName,
    a.UserId,
    ISNULL(d.Name, '')              AS PositionName,
    a.DisplayName
FROM UserProfiles a
LEFT JOIN AspNetUsers b          ON a.UserId = b.Id
LEFT JOIN PositionMasters c      ON a.CompCd = c.CompCd AND a.PositionId = c.Id
LEFT JOIN PositionMasterLoc d    ON c.Id = d.PositionId AND d.LangCode = @LangCode
LEFT JOIN DepartmentMasters e    ON a.CompCd = e.CompCd AND a.DepartmentId = e.Id
LEFT JOIN DepartmentMasterLoc f  ON e.Id = f.DepartmentId AND f.LangCode = @LangCode
LEFT JOIN CompMasters g          ON a.CompCd = g.CompCd
ORDER BY a.CompCd, e.SortOrder, c.RankLevel, a.DisplayName;";
                cmd.Parameters.Add(new SqlParameter("@LangCode", SqlDbType.NVarChar, 8) { Value = langCode });

                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    var compCd = rd["CompCd"] as string ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(compCd))
                        continue;

                    var compName = rd["CompName"] as string ?? compCd;

                    var deptIdObj = rd["DepartmentId"];
                    if (deptIdObj == null || deptIdObj == DBNull.Value)
                        continue;

                    var deptId = Convert.ToString(deptIdObj) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(deptId))
                        continue;

                    var deptName = rd["DeptName"] as string ?? deptId;

                    var userId = rd["UserId"] as string ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(userId))
                        continue;

                    var positionName = rd["PositionName"] as string ?? string.Empty;
                    var displayName = rd["DisplayName"] as string ?? string.Empty;

                    // 회사 노드
                    if (!compMap.TryGetValue(compCd, out var compNode))
                    {
                        compNode = new OrgTreeNode
                        {
                            NodeId = compCd,
                            Name = compName,
                            NodeType = "Branch",
                            ParentId = null
                        };
                        compMap[compCd] = compNode;
                        orgNodes.Add(compNode);
                    }

                    // 부서 노드 (유니크 NodeId: compCd:deptId)
                    var deptKey = compCd + ":" + deptId;
                    if (!deptMap.TryGetValue(deptKey, out var deptNode))
                    {
                        deptNode = new OrgTreeNode
                        {
                            NodeId = deptKey,
                            Name = deptName,
                            NodeType = "Dept",
                            ParentId = compNode.NodeId
                        };
                        deptMap[deptKey] = deptNode;
                        compNode.Children.Add(deptNode);
                    }

                    // 사용자 노드
                    var caption = string.IsNullOrWhiteSpace(positionName)
                        ? displayName
                        : positionName + " " + displayName;

                    var userNode = new OrgTreeNode
                    {
                        NodeId = userId,
                        Name = caption,
                        NodeType = "User",
                        ParentId = deptNode.NodeId
                    };

                    deptNode.Children.Add(userNode);
                }
            }
            catch
            {
                // 실패 시 빈 리스트 반환 (기존 동작 유지)
            }

            return orgNodes;
        }
        private string ToContentRootAbsolute(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            var normalized = path
                .Replace('/', Path.DirectorySeparatorChar)
                .Replace('\\', Path.DirectorySeparatorChar);

            if (Path.IsPathRooted(normalized))
                return normalized;

            return Path.Combine(_env.ContentRootPath, normalized);
        }

        // ContentRoot 기준 상대 경로(절대 → 상대) 변환
        private string ToContentRootRelative(string? fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath)) return string.Empty;

            var root = _env.ContentRootPath.TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar);

            var normalized = fullPath
                .Replace('/', Path.DirectorySeparatorChar)
                .Replace('\\', Path.DirectorySeparatorChar);

            if (Path.IsPathRooted(normalized) &&
                normalized.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                return Path.GetRelativePath(root, normalized);
            }

            return normalized;
        }

        // ========== New ==========
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

        // ========== Create (GET) ==========
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

        // ========== CSRF ==========
        [HttpGet("Csrf")]
        [Produces("application/json")]
        public IActionResult Csrf()
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

        private async Task InsertDocumentFileRowAsync(
            string docId,
            string fileKey,
            string originalName,
            string storagePath,
            string contentType,
            long byteSize,
            byte[] sha256Bytes,
            string uploadedBy
        )
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
        // ========== Create (POST) ==========
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
            //    템플릿 결재 JSON이 비면 toEmails 로 A1 A2 ... 생성
            // ----------------------------
            var approvalsJson = ExtractApprovalsJson(normalizedDesc);

            if (string.IsNullOrWhiteSpace(approvalsJson) || approvalsJson.Trim() == "[]")
            {
                if (toEmails.Count > 0)
                {
                    var built = new List<Dictionary<string, object>>();

                    for (int i = 0; i < toEmails.Count; i++)
                    {
                        var email = (toEmails[i] ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(email)) continue;

                        built.Add(new Dictionary<string, object>
                        {
                            ["roleKey"] = $"A{i + 1}",
                            ["approverType"] = "Person",
                            ["required"] = true,
                            ["value"] = email
                        });
                    }

                    if (built.Count == 0)
                        return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "approvals", detail = "No valid approver emails." });

                    approvalsJson = System.Text.Json.JsonSerializer.Serialize(built);
                }
                else
                {
                    // 결재자 해석 자체가 실패한 상태면 저장을 막는 것이 안전
                    return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "approvals", detail = "Approver not resolved (toEmails empty)." });
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

            object? mailInfo = null;
            //try
            //{
            //    mailInfo = await SendSubmitMailAsync(docId, dto, normalizedDesc, toEmails, title, tc);
            //}
            //catch (Exception ex)
            //{
            //    _log.LogWarning(ex, "SendSubmitMailAsync failed docId={docId}", docId);
            //    mailInfo = new { from = "", to = Array.Empty<string>(), cc = Array.Empty<string>(), bcc = Array.Empty<string>(), subject = "", body = "", sent = false, sentCount = 0, error = ex.Message };
            //}

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

        #region Mail (extracted)

        private async Task<object?> SendSubmitMailAsync(
            string docId,
            ComposePostDto dto,
            string normalizedDesc,
            List<string> toEmails,
            string? title,
            string tc
        )
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

        #endregion

        // ========== 승인 처리 ==========
        public sealed class ApproveDto
        {
            public string? docId { get; set; }
            public string? action { get; set; }
            public int slot { get; set; } = 1;
            public string? comment { get; set; }
        }

        public sealed class DocCommentPostDto
        {
            public string? docId { get; set; }
            public int? parentCommentId { get; set; }
            public string? body { get; set; }

            // 업로드 메타 (JS에서 files: [...] 로 보내는 구조와 맞추기)
            public List<CommentFileDto>? files { get; set; }
        }

        public sealed class CommentFileDto
        {
            public string? fileKey { get; set; }
            public string? originalName { get; set; }
            public string? contentType { get; set; }
            public long byteSize { get; set; }
        }

        private async Task<(bool canRecall, bool canApprove, bool canHold, bool canReject)>
    GetApprovalCapabilitiesAsync(string docId)
        {
            var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(userId))
                return (false, false, false, false);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            string? status = null;
            string? createdBy = null;

            // 1) 문서 상태 / 작성자 조회
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT TOP 1 Status, CreatedBy
FROM dbo.Documents
WHERE DocId = @id;";
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });

                await using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    status = r["Status"] as string ?? string.Empty;
                    createdBy = r["CreatedBy"] as string ?? string.Empty;
                }
            }

            if (status == null)
                return (false, false, false, false);

            var st = status ?? string.Empty;
            var isCreator = string.Equals(createdBy ?? string.Empty, userId, StringComparison.OrdinalIgnoreCase);
            var isRecalled = st.Equals("Recalled", StringComparison.OrdinalIgnoreCase);
            var isFinal =
                st.StartsWith("Approved", StringComparison.OrdinalIgnoreCase) ||
                st.StartsWith("Rejected", StringComparison.OrdinalIgnoreCase);

            // 회수: 작성자 && 아직 Pending 상태일 때만 허용
            var canRecall = isCreator && st.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);

            // 2) 현재 Pending 단계 찾기
            int currentStep = 0;
            await using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = @"
SELECT MIN(StepOrder)
FROM dbo.DocumentApprovals
WHERE DocId = @id
  AND ISNULL(Status, N'Pending') LIKE N'Pending%';";
                cmd2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });

                var obj = await cmd2.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                    currentStep = Convert.ToInt32(obj);
            }

            if (currentStep == 0 || isFinal || isRecalled)
                return (canRecall, false, false, false);

            // 3) 해당 단계 승인자(UserId, Status) 조회
            string? stepUserId = null;
            string stepStatus = string.Empty;

            await using (var cmd3 = conn.CreateCommand())
            {
                cmd3.CommandText = @"
SELECT TOP 1 UserId, ISNULL(Status, N'Pending') AS Status
FROM dbo.DocumentApprovals
WHERE DocId = @id
  AND StepOrder = @step;";
                cmd3.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                cmd3.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = currentStep });

                await using var r = await cmd3.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    stepUserId = r["UserId"] as string ?? string.Empty;
                    stepStatus = r["Status"] as string ?? string.Empty;
                }
            }

            if (string.IsNullOrWhiteSpace(stepUserId))
                return (canRecall, false, false, false);

            var isMyStep = string.Equals(stepUserId, userId, StringComparison.OrdinalIgnoreCase);
            var isPending = stepStatus.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);

            var canDo = isMyStep && isPending && !isRecalled && !isFinal;

            return (canRecall, canDo, canDo, canDo);
        }

        [HttpPost("ApproveOrHold")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        // 2025.12.09 Changed: n차 승인 메일도 BuildSubmissionMail 사용하여 최초 상신 메일과 동일한 제목과 작성자 이름을 유지하고 수신자 표시 이름만 교체하도록 수정
        public async Task<IActionResult> ApproveOrHold([FromBody] ApproveDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.docId) || string.IsNullOrWhiteSpace(dto.action))
                return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" } });

            var actionLower = dto.action!.ToLowerInvariant();

            // ===== 1) 회수 전용 처리 =====
            if (actionLower == "recall")
            {
                var csRecall = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using (var conn = new SqlConnection(csRecall))
                {
                    await conn.OpenAsync();

                    // 1-1) 문서 상태를 Recalled 로 변경 (현재 Pending 단계에서만 허용)
                    await using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
UPDATE dbo.Documents
SET Status = N'Recalled'
WHERE DocId = @DocId
  AND ISNULL(Status,N'') LIKE N'Pending%';";
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        await cmd.ExecuteNonQueryAsync();
                    }

                    // 1-2) 승인선은 Status 를 그대로 유지하고, 회수 이력만 남김
                    await using (var ac = conn.CreateCommand())
                    {
                        ac.CommandText = @"
UPDATE dbo.DocumentApprovals
SET Action    = N'Recalled',
    ActedAt   = SYSUTCDATETIME(),
    ActorName = @actor
WHERE DocId = @DocId
  AND ISNULL(Status,N'Pending') LIKE N'Pending%';";
                        ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        ac.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200)
                        {
                            Value = GetCurrentUserDisplayNameStrict() ?? string.Empty
                        });
                        await ac.ExecuteNonQueryAsync();
                    }
                }

                // 클라이언트에는 문서 최종 상태만 전달
                return Json(new { ok = true, docId = dto.docId, status = "Recalled" });
            }

            // ===== 2) 승인 보류 반려 공통 처리 =====
            //var outXlsx = ResolveOutputPathFromDocId(dto.docId);
            //if (string.IsNullOrWhiteSpace(outXlsx) || !System.IO.File.Exists(outXlsx))
            //    return NotFound(new { messages = new[] { "DOC_Err_DocumentNotFound" } });
            string outXlsx = string.Empty;

            try
            {
                var csOut = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using (var connOut = new SqlConnection(csOut))
                {
                    await connOut.OpenAsync();

                    await using var cmdOut = connOut.CreateCommand();
                    cmdOut.CommandText = @"
SELECT TOP (1) OutputPath
FROM dbo.Documents
WHERE DocId = @DocId;";
                    cmdOut.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });

                    var obj = await cmdOut.ExecuteScalarAsync();
                    outXlsx = (obj as string) ?? string.Empty;
                }

                // DB에 상대 경로(App_Data\Docs\... 형태)로 저장되어 있는 경우 ContentRoot 기준으로 보정
                if (!string.IsNullOrWhiteSpace(outXlsx) && !Path.IsPathRooted(outXlsx))
                {
                    outXlsx = Path.Combine(
                        _env.ContentRootPath,
                        outXlsx.TrimStart('\\', '/')
                    );
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold: output path resolve failed docId={docId}", dto.docId);
                outXlsx = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(outXlsx) || !System.IO.File.Exists(outXlsx))
                return NotFound(new { messages = new[] { "DOC_Err_DocumentNotFound" } });

            var approverId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            var approverName = GetCurrentUserDisplayNameStrict();

            string? approverDisplayText = null;
            string? signatureRelativePath = null;

            // --- 현재 로그인 사용자의 프로필 스냅샷 (결재 캡션 서명용) ---
            try
            {
                var csProfile = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using (var connProfile = new SqlConnection(csProfile))
                {
                    await connProfile.OpenAsync();

                    await using var up = connProfile.CreateCommand();
                    up.CommandText = @"
SELECT TOP (1)
       u.DisplayName,
       COALESCE(pl.Name, pm.Name, N'') AS PositionName,
       COALESCE(dl.Name, dm.Name, N'') AS DepartmentName,
       u.SignatureRelativePath AS SignaturePath
FROM dbo.UserProfiles AS u
LEFT JOIN dbo.DepartmentMasters     AS dm ON dm.CompCd = u.CompCd AND dm.Id       = u.DepartmentId
LEFT JOIN dbo.DepartmentMasterLoc   AS dl ON dl.DepartmentId = dm.Id AND dl.LangCode = N'ko'
LEFT JOIN dbo.PositionMasters       AS pm ON pm.CompCd = u.CompCd AND pm.Id       = u.PositionId
LEFT JOIN dbo.PositionMasterLoc     AS pl ON pl.PositionId   = pm.Id AND pl.LangCode = N'ko'
WHERE u.UserId = @uid;";
                    up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });

                    await using var r = await up.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        var disp = r["DisplayName"] as string ?? string.Empty;
                        var pos = r["PositionName"] as string ?? string.Empty;
                        var dept = r["DepartmentName"] as string ?? string.Empty;
                        var sig = r["SignaturePath"] as string ?? string.Empty;

                        approverDisplayText = string.Join(" ",
                            new[] { dept, pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));

                        signatureRelativePath = string.IsNullOrWhiteSpace(sig) ? null : sig;
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold: 프로필 스냅샷 조회 실패 docId={docId}", dto.docId);
            }

            string newStatus = "Updated";

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                // --- 2 1) 현재 로그인 사용자가 담당인 Pending 단계 StepOrder 계산 ---
                int? currentStep = null;
                await using (var findStep = conn.CreateCommand())
                {
                    findStep.CommandText = @"
SELECT TOP (1) StepOrder
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND ISNULL(Status,N'Pending') LIKE N'Pending%'
  AND (
        UserId = @UserId
        OR ApproverValue = @UserId
      )
ORDER BY StepOrder;";
                    findStep.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    findStep.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId });

                    var stepObj = await findStep.ExecuteScalarAsync();
                    if (stepObj != null && stepObj != DBNull.Value)
                        currentStep = Convert.ToInt32(stepObj);
                }

                // 현재 사용자가 결재할 수 있는 단계가 없으면 차단
                if (currentStep == null)
                    return Forbid();

                var step = currentStep.Value;

                // --- 2 2) 현재 단계 DocumentApprovals 업데이트 ---
                await using (var u = conn.CreateCommand())
                {
                    u.CommandText = @"
UPDATE dbo.DocumentApprovals
SET Action              = @act,
    ActedAt             = SYSUTCDATETIME(),
    ActorName           = @actor,
    Status              = CASE 
                             WHEN @act = N'approve' THEN N'Approved'
                             WHEN @act = N'hold'    THEN N'OnHold'
                             WHEN @act = N'reject'  THEN N'Rejected'
                             ELSE ISNULL(Status, N'Updated')
                          END,
    UserId              = COALESCE(UserId, @uid),
    ApproverDisplayText = COALESCE(@displayText, ApproverDisplayText),
    SignaturePath       = COALESCE(@sigPath, SignaturePath)
WHERE DocId = @id AND StepOrder = @step;";
                    u.Parameters.Add(new SqlParameter("@act", SqlDbType.NVarChar, 20) { Value = actionLower });
                    u.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName ?? string.Empty });
                    u.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    u.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = step });
                    u.Parameters.Add(new SqlParameter("@displayText", SqlDbType.NVarChar, 400)
                    {
                        Value = (object?)approverDisplayText ?? DBNull.Value
                    });
                    u.Parameters.Add(new SqlParameter("@sigPath", SqlDbType.NVarChar, 400)
                    {
                        Value = (object?)signatureRelativePath ?? DBNull.Value
                    });

                    await u.ExecuteNonQueryAsync();
                }

                // --- 2 3) Documents.Status 1차 업데이트 (현재 단계 기준) ---
                newStatus = actionLower switch
                {
                    "approve" => $"ApprovedA{step}",
                    "hold" => $"OnHoldA{step}",
                    "reject" => $"RejectedA{step}",
                    _ => "Updated"
                };

                await using (var u = conn.CreateCommand())
                {
                    u.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                    u.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = newStatus });
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    await u.ExecuteNonQueryAsync();
                }

                // --- 2 4) 승인일 때만 다음 단계 처리 및 메일 발송 ---
                if (actionLower == "approve")
                {
                    var next = step + 1;

                    var toList = await GetNextApproverEmailsFromDbAsync(conn, dto.docId, next);

                    if (toList.Count == 0)
                    {
                        await using var up2a = conn.CreateCommand();
                        up2a.CommandText = @"UPDATE dbo.Documents SET Status = N'Approved' WHERE DocId = @id;";
                        up2a.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        await up2a.ExecuteNonQueryAsync();
                        newStatus = "Approved";
                    }
                    else
                    {
                        // 1) 문서 기본 정보 조회 (작성자 이름 문서 제목 작성 시각)
                        string? mailAuthor = null;
                        string? mailDocTitle = null;
                        DateTime? createdAtLocalForMail = null;

                        await using (var q = conn.CreateCommand())
                        {
                            q.CommandText = @"
SELECT TOP 1
       -- 작성자 이름은 UserProfiles.DisplayName 을 우선 사용
       ISNULL(up.DisplayName,
              ISNULL(d.CreatedByName, d.CreatedBy)) AS AuthorName,
       ISNULL(d.TemplateTitle, N'')                 AS TemplateTitle,
       d.CreatedAt                                  AS CreatedAtUtc
FROM dbo.Documents      AS d
LEFT JOIN dbo.UserProfiles AS up
       ON up.UserId = d.CreatedBy
WHERE d.DocId = @id;";
                            q.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });

                            await using var r = await q.ExecuteReaderAsync();
                            if (await r.ReadAsync())
                            {
                                mailAuthor = r["AuthorName"] as string ?? string.Empty;
                                mailDocTitle = r["TemplateTitle"] as string ?? string.Empty;

                                if (r["CreatedAtUtc"] != DBNull.Value)
                                {
                                    var createdAtUtc = (DateTime)r["CreatedAtUtc"];
                                    if (createdAtUtc.Kind != DateTimeKind.Utc)
                                        createdAtUtc = DateTime.SpecifyKind(createdAtUtc, DateTimeKind.Utc);

                                    var tz = ResolveCompanyTimeZone();
                                    createdAtLocalForMail = TimeZoneInfo.ConvertTimeFromUtc(createdAtUtc, tz);
                                }
                            }
                        }

                        if (string.IsNullOrWhiteSpace(mailAuthor))
                            mailAuthor = approverName ?? string.Empty;

                        if (string.IsNullOrWhiteSpace(mailDocTitle))
                            mailDocTitle = dto.docId;

                        // 2) 다음 단계 상태를 PendingA{next} 로 변경
                        await using (var up2 = conn.CreateCommand())
                        {
                            up2.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                            up2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = $"PendingA{next}" });
                            up2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            await up2.ExecuteNonQueryAsync();
                        }
                        newStatus = $"PendingA{next}";

                        // 3) 중복 제거된 수신자 목록
                        var distinctTo = toList
                            .Where(e => !string.IsNullOrWhiteSpace(e))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        // 4) 상신 메일 포맷 그대로 사용하여 개별 발송
                        // 사장님 지시로 메일 발송 금지
                        //foreach (var to in distinctTo)
                        //{
                        //    try
                        //    {
                        //        // 수신자 인사말용 표시 이름 (예: 생산차장)
                        //        var recipientDisplay = GetDisplayNameByEmailStrict(to);
                        //        if (string.IsNullOrWhiteSpace(recipientDisplay))
                        //            recipientDisplay = FallbackNameFromEmail(to);

                        //        var mail = BuildSubmissionMail(
                        //            mailAuthor!,
                        //            mailDocTitle!,
                        //            dto.docId,
                        //            recipientDisplay,      // "안녕하세요, {0}님" 에 들어갈 이름
                        //            createdAtLocalForMail  // 최초 작성 시각(로컬)
                        //        );

                        //        await _emailSender.SendEmailAsync(to, mail.subject, mail.bodyHtml);

                        //        _log.LogInformation(
                        //            "NextApproverMail SEND OK DocId {DocId} NextStep {NextStep} To {To}",
                        //            dto.docId, next, to);
                        //    }
                        //    catch (Exception exMail)
                        //    {
                        //        _log.LogError(
                        //            exMail,
                        //            "NextApproverMail SEND ERROR DocId {DocId} NextStep {NextStep} To {To}",
                        //            dto.docId, next, to);
                        //    }
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold flow failed docId={docId}", dto.docId);
            }

            var previewJson = BuildPreviewJsonFromExcel(outXlsx!);
            return Json(new { ok = true, docId = dto.docId, status = newStatus, previewJson });
        }

        // 2025.12.08 Changed: 트랜잭션 없는 호출을 위해 tx 를 nullable 로 변경하고 옵션 적용
        private static async Task<string?> TryResolveUserIdAsync(SqlConnection conn, SqlTransaction? tx, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            await using var cmd = conn.CreateCommand();
            if (tx != null) cmd.Transaction = tx;
            cmd.CommandText = @"
SELECT TOP 1 Id
FROM dbo.AspNetUsers
WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);";
            cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            var id = (string?)await cmd.ExecuteScalarAsync();
            if (!string.IsNullOrWhiteSpace(id)) return id;

            await using var cmd2 = conn.CreateCommand();
            if (tx != null) cmd2.Transaction = tx;
            cmd2.CommandText = @"
SELECT TOP 1 COALESCE(p.UserId, u.Id)
FROM dbo.UserProfiles p
LEFT JOIN dbo.AspNetUsers u ON u.Id = p.UserId
WHERE p.UserId=@v OR p.Email=@v OR p.DisplayName=@v OR p.Name=@v;";
            cmd2.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            id = (string?)await cmd2.ExecuteScalarAsync();
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        // ★ 추가: 문서의 TemplateVersionId + DocTemplateApproval.Slot 으로
        //   승인란(H4 / I4 / J4 등) 좌표를 찾아오는 헬퍼입니다.
        //   dto.slot(=StepOrder) 값을 넘겨서 호출하게 됩니다.
        private static async Task<(int sheet, int row, int col)?> GetApprovalCellAsync(
            SqlConnection conn,
            SqlTransaction? tx,
            string docId,
            int stepOrder)
        {
            await using var cmd = conn.CreateCommand();
            if (tx != null) cmd.Transaction = tx;

            cmd.CommandText = @"
SELECT TOP (1) ta.Sheet,
               ta.CellRow,
               ta.CellColumn
FROM dbo.Documents          AS d
JOIN dbo.DocTemplateVersion AS tv ON tv.Id = d.TemplateVersionId
JOIN dbo.DocTemplateApproval AS ta ON ta.VersionId = tv.Id
WHERE d.DocId = @DocId
  AND ta.Slot = @StepOrder;";   // Slot = 현재 승인 단계(1차/2차/3차)

            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = stepOrder });

            await using var r = await cmd.ExecuteReaderAsync();
            if (await r.ReadAsync())
            {
                var sheet = Convert.ToInt32(r["Sheet"]);
                var row = Convert.ToInt32(r["CellRow"]);
                var col = Convert.ToInt32(r["CellColumn"]);
                return (sheet, row, col);
            }

            return null;
        }


        /// 현재 단계 이후(nextStep)의 승인자를 DocumentApprovals 에서 읽어서
        /// UserId 를 보정하고 메일 주소 목록을 반환합니다.
        private static async Task<List<string>> GetNextApproverEmailsFromDbAsync(
    SqlConnection conn,
    string docId,
    int stepOrder)
        {
            var result = new List<string>();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT ApproverValue
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder = @StepOrder
  AND ISNULL(Status, N'Pending') LIKE N'Pending%'
  AND ApproverValue IS NOT NULL
  AND LTRIM(RTRIM(ApproverValue)) <> N'';";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = stepOrder });

            await using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                var v = r[0] as string;
                if (!string.IsNullOrWhiteSpace(v))
                    result.Add(v.Trim());
            }

            return result
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private async Task<IList<object>> LoadApprovalsForDetailAsync(string docId)
        {
            var list = new List<object>();

            var cs = _cfg.GetConnectionString("DefaultConnection");
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            const string sql = @"
SELECT  a.DocId,
        a.StepOrder,
        a.RoleKey,
        a.ApproverValue,
        a.UserId,
        a.Status,
        a.Action,
        a.ActedAt,
        a.ActorName,
        a.ApproverDisplayText,
        up.DisplayName AS ApproverName
FROM    dbo.DocumentApprovals a
LEFT JOIN dbo.UserProfiles up
       ON up.UserId = a.UserId
WHERE   a.DocId = @DocId
ORDER BY a.StepOrder;";

            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

            await using var rdr = await cmd.ExecuteReaderAsync();
            while (await rdr.ReadAsync())
            {
                var stepOrder = rdr["StepOrder"] is int so ? so : Convert.ToInt32(rdr["StepOrder"]);
                var roleKey = rdr["RoleKey"]?.ToString() ?? string.Empty;
                var approverValue = rdr["ApproverValue"]?.ToString() ?? string.Empty;   // 이메일 등
                var approverName = rdr["ApproverName"]?.ToString() ?? string.Empty;    // UserProfiles.DisplayName
                var storedDisplay = rdr["ApproverDisplayText"]?.ToString() ?? string.Empty;
                var status = rdr["Status"]?.ToString() ?? string.Empty;
                var action = rdr["Action"]?.ToString() ?? string.Empty;
                var actedAtObj = rdr["ActedAt"];
                DateTime? actedAt = actedAtObj is DateTime dt ? dt : null;

                // 담당자 텍스트: 이름 > 기존 ApproverDisplayText > 이메일
                var displayText =
                    !string.IsNullOrWhiteSpace(approverName) ? approverName :
                    !string.IsNullOrWhiteSpace(storedDisplay) ? storedDisplay :
                    approverValue;

                list.Add(new
                {
                    step = stepOrder,
                    roleKey = roleKey,
                    approverValue = approverValue,
                    approverName = approverName,
                    approverDisplayText = displayText,
                    status = status,
                    action = action,
                    actedAt = actedAt,
                    actedAtText = actedAt.HasValue
                                           ? ToLocalStringFromUtc(actedAt.Value)
                                           : string.Empty,
                    actorName = rdr["ActorName"]?.ToString() ?? string.Empty
                });
            }

            return list;
        }

        [HttpGet("Detail/{id?}")]
        // 2025.12.02 Changed: Detail 에서 템플릿 버전별 결재 셀 A1 정보를 조회해 ViewBag.ApprovalCellsJson 으로 전달
        public async Task<IActionResult> Detail(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                var qid = HttpContext?.Request?.Query["id"].ToString();
                if (!string.IsNullOrWhiteSpace(qid))
                    id = qid;
            }

            if (string.IsNullOrWhiteSpace(id))
                return RedirectToAction("Board");

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // ---------------- 1) 기본 문서 정보 ----------------
            string? templateCode = null;
            string? templateTitle = null;
            string? status = null;
            string? descriptorJson = null;
            string? outputPath = null;
            string? compCd = null;

            string? creatorUserId = null;
            string? creatorNameRaw = null;

            DateTime? createdAtUtc = null;
            DateTime? updatedAtUtc = null;

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT TOP 1
       d.DocId,
       d.TemplateCode,
       d.TemplateTitle,
       d.Status,
       d.DescriptorJson,
       d.OutputPath,
       d.CompCd,
       d.CreatedBy,
       d.CreatedByName,
       d.CreatedAt,
       d.UpdatedAt,
       COALESCE(p.DisplayName, d.CreatedByName, u.UserName, u.Email) AS CreatorDisplayName
FROM dbo.Documents d
LEFT JOIN dbo.AspNetUsers  u ON u.Id     = d.CreatedBy
LEFT JOIN dbo.UserProfiles p ON p.UserId = d.CreatedBy
WHERE d.DocId = @DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var rd = await cmd.ExecuteReaderAsync();
                if (!await rd.ReadAsync())
                {
                    ViewBag.Message = _S["DOC_Err_DocumentNotFound"];
                    ViewBag.DocId = id;
                    ViewBag.DocumentId = id;
                    return View("Detail");
                }

                templateCode = rd["TemplateCode"] as string ?? "";
                templateTitle = rd["TemplateTitle"] as string ?? "";
                status = rd["Status"] as string ?? "";
                descriptorJson = rd["DescriptorJson"] as string ?? "{}";
                outputPath = rd["OutputPath"] as string ?? "";
                compCd = rd["CompCd"] as string ?? "";

                creatorUserId = rd["CreatedBy"] as string ?? "";
                creatorNameRaw = rd["CreatorDisplayName"] as string ?? "";

                createdAtUtc = rd["CreatedAt"] is DateTime cdt ? cdt : (DateTime?)null;
                updatedAtUtc = rd["UpdatedAt"] is DateTime udt ? udt : (DateTime?)null;
            }

            // 작성자 DisplayName 우선 해석
            string? creatorDisplayName = creatorNameRaw;

            if (!string.IsNullOrWhiteSpace(creatorUserId))
            {
                await using var pc = conn.CreateCommand();
                pc.CommandText = @"
SELECT TOP 1 DisplayName
FROM dbo.UserProfiles
WHERE UserId = @UserId;";
                pc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = creatorUserId });

                var objName = await pc.ExecuteScalarAsync();
                if (objName != null && objName != DBNull.Value)
                {
                    creatorDisplayName = Convert.ToString(objName) ?? creatorDisplayName;
                }
            }

            // ---------------- 1-0) 작성/회수 시각(표시용) ----------------
            DateTime? createdAtForUi = null;
            string createdAtText = string.Empty;

            if (createdAtUtc.HasValue)
            {
                createdAtForUi = DateTime.SpecifyKind(createdAtUtc.Value, DateTimeKind.Utc);
                createdAtText = ToLocalStringFromUtc(createdAtForUi.Value);
            }

            DateTime? recalledAtForUi = null;
            string recalledAtText = string.Empty;

            try
            {
                await using var rcCmd = conn.CreateCommand();
                rcCmd.CommandText = @"
SELECT TOP 1 a.ActedAt
FROM dbo.DocumentApprovals a
WHERE a.DocId = @DocId
  AND (
        UPPER(ISNULL(a.Action,'')) = 'RECALLED'
     OR UPPER(ISNULL(a.Status,'')) = 'RECALLED'
  )
  AND a.ActedAt IS NOT NULL
ORDER BY a.ActedAt DESC;";
                rcCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                var obj = await rcCmd.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                {
                    var dt = (obj is DateTime dtx) ? dtx : Convert.ToDateTime(obj);
                    recalledAtForUi = DateTime.SpecifyKind(dt, DateTimeKind.Utc);
                    recalledAtText = ToLocalStringFromUtc(recalledAtForUi.Value);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load recalled ActedAt failed for {DocId}", id);
            }

            // ---------------- 1-1) TemplateCode -> 템플릿 버전 Id ----------------
            int? templateVersionId = null;
            if (!string.IsNullOrWhiteSpace(templateCode))
            {
                await using var tvCmd = conn.CreateCommand();
                tvCmd.CommandText = @"
SELECT TOP 1 v.Id
FROM dbo.DocTemplateVersion v
INNER JOIN dbo.DocTemplateMaster m ON v.TemplateId = m.Id
WHERE m.DocCode = @TemplateCode
ORDER BY v.VersionNo DESC, v.Id DESC;";
                tvCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 50) { Value = templateCode });

                var obj = await tvCmd.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                    templateVersionId = Convert.ToInt32(obj);
            }

            // ---------------- 2) 프리뷰 JSON ----------------
            var inputsJson = "{}";
            string previewJson = "{}";

            try
            {
                var raw = outputPath ?? string.Empty;
                var resolved = raw;

                if (!string.IsNullOrWhiteSpace(resolved))
                {
                    if (!System.IO.Path.IsPathRooted(resolved))
                    {
                        resolved = resolved.TrimStart('\\', '/');
                        resolved = System.IO.Path.Combine(_env.ContentRootPath, resolved);
                    }

                    resolved = resolved.Replace('/', System.IO.Path.DirectorySeparatorChar)
                                       .Replace('\\', System.IO.Path.DirectorySeparatorChar);
                }

                var exists = !string.IsNullOrWhiteSpace(resolved) && System.IO.File.Exists(resolved);
                if (!exists)
                {
                    _log.LogWarning(
                        "Detail: output excel not found. DocId={DocId}, OutputPath(raw)={Raw}, OutputPath(resolved)={Resolved}, CWD={Cwd}, ContentRoot={Root}",
                        id, raw, resolved, System.IO.Directory.GetCurrentDirectory(), _env.ContentRootPath
                    );
                }
                else
                {
                    previewJson = BuildPreviewJsonFromExcel(resolved);
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Detail: BuildPreviewJsonFromExcel failed for {DocId}. OutputPath={OutputPath}", id, outputPath ?? string.Empty);
                previewJson = "{}";
            }

            // ---------------- 3) 승인 정보 DisplayName 조인 ----------------
            var approvals = new List<object>();
            await using (var ac = conn.CreateCommand())
            {
                ac.CommandText = @"
SELECT  a.StepOrder,
        a.RoleKey,
        a.ApproverValue,
        a.UserId,
        a.Status,
        a.Action,
        a.ActedAt,
        a.ActorName,
        COALESCE(a.ApproverDisplayText, up.DisplayName, a.ActorName, a.ApproverValue) AS ApproverDisplayText,
        a.SignaturePath
FROM dbo.DocumentApprovals a
LEFT JOIN dbo.UserProfiles up
       ON up.UserId = a.UserId
WHERE a.DocId = @DocId
ORDER BY a.StepOrder;";
                ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    DateTime? acted = r.IsDBNull(6) ? (DateTime?)null : r.GetDateTime(6);
                    string? when = acted.HasValue
                        ? ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                        : null;

                    var approverDisplayText = r["ApproverDisplayText"] as string ?? string.Empty;

                    approvals.Add(new
                    {
                        step = r.GetInt32(0),
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue = r["ApproverValue"]?.ToString(),
                        userId = r["UserId"]?.ToString(),
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = when,
                        actorName = r["ActorName"]?.ToString(),
                        approverDisplayText,
                        signaturePath = r["SignaturePath"]?.ToString()
                    });
                }
            }

            // ---------------- 3-1) 템플릿 결재 셀 A1 매핑 ----------------
            var approvalCells = new List<object>();
            if (templateVersionId.HasValue)
            {
                await using var ac2 = conn.CreateCommand();
                ac2.CommandText = @"
SELECT Slot,
       ISNULL(CellA1, A1) AS A1
FROM dbo.DocTemplateApproval
WHERE VersionId = @VersionId
ORDER BY Slot;";
                ac2.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.Int) { Value = templateVersionId.Value });

                await using var rc = await ac2.ExecuteReaderAsync();
                while (await rc.ReadAsync())
                {
                    approvalCells.Add(new
                    {
                        Step = rc.GetInt32(0),
                        A1 = rc["A1"] as string ?? string.Empty
                    });
                }
            }

            // ---------------- 3-2) 문서 첨부 파일 목록(DocumentFiles) ----------------
            var docFiles = new List<object>();
            try
            {
                await using var fc = conn.CreateCommand();
                fc.CommandText = @"
SELECT FileKey,
       OriginalName,
       StoragePath,
       ContentType,
       ByteSize,
       Sha256,
       UploadedBy,
       UploadedAt
FROM dbo.DocumentFiles
WHERE DocId = @DocId
ORDER BY UploadedAt, FileId ASC;";
                fc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var fr = await fc.ExecuteReaderAsync();
                while (await fr.ReadAsync())
                {
                    var fileKey = fr["FileKey"] as string ?? string.Empty;
                    var originalName = fr["OriginalName"] as string ?? string.Empty;
                    var contentType = fr["ContentType"] as string ?? string.Empty;
                    long byteSize = 0;
                    if (fr["ByteSize"] != DBNull.Value) byteSize = Convert.ToInt64(fr["ByteSize"]);

                    DateTime? upUtc = fr["UploadedAt"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(fr["UploadedAt"]);
                    string? upText = null;
                    if (upUtc.HasValue)
                    {
                        var up = DateTime.SpecifyKind(upUtc.Value, DateTimeKind.Utc);
                        upText = ToLocalStringFromUtc(up);
                    }

                    docFiles.Add(new
                    {
                        fileKey,
                        originalName,
                        contentType,
                        byteSize,
                        uploadedAtText = upText
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load DocumentFiles failed for {DocId}", id);
            }

            // ---------------- 3-3) 공유 대상자 목록(DocumentShares) ----------------
            var sharedRecipientUserIds = new List<string>();
            try
            {
                await using var sc = conn.CreateCommand();
                sc.CommandText = @"
SELECT s.UserId
FROM dbo.DocumentShares s
WHERE s.DocId = @DocId
  AND ISNULL(s.IsRevoked, 0) = 0
  AND (s.ExpireAt IS NULL OR s.ExpireAt > GETUTCDATE())
ORDER BY s.CreatedAt ASC, s.Id ASC;";
                sc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var sr = await sc.ExecuteReaderAsync();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                while (await sr.ReadAsync())
                {
                    var uid = sr.IsDBNull(0) ? null : sr.GetString(0);
                    if (string.IsNullOrWhiteSpace(uid)) continue;
                    if (seen.Add(uid))
                        sharedRecipientUserIds.Add(uid);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load DocumentShares failed for {DocId}", id);
            }

            // ---------------- 3-4) (수정) 공유 대상자 콤보 후보 데이터(OrgTreeNodes) 로드 ----------------
            // 공통 헬퍼(BuildOrgTreeNodesAsync)로 통일 (Create(GET)과 동일 로직 재사용)
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
                _log.LogWarning(ex, "Detail: load OrgTreeNodes failed for {DocId}, CompCd={CompCd}", id, compCd ?? string.Empty);
                ViewBag.OrgTreeNodes = Array.Empty<OrgTreeNode>();
            }

            // ---------------- 4) 조회 로그 기록 ----------------
            try
            {
                await LogDocumentViewAsync(conn, id);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: LogDocumentViewAsync failed for {DocId}", id);
            }

            // ---------------- 5) 조회 로그 조회 ----------------
            var viewLogs = new List<object>();
            try
            {
                await using var vc = conn.CreateCommand();
                vc.CommandText = @"
SELECT TOP 200
       v.ViewedAt,
       v.ViewerId,
       v.ViewerRole,
       v.ClientIp,
       v.UserAgent,
       COALESCE(p.DisplayName, u.UserName, u.Email) AS UserName
FROM dbo.DocumentViewLogs v
LEFT JOIN dbo.UserProfiles p ON p.UserId = v.ViewerId
LEFT JOIN dbo.AspNetUsers  u ON u.Id     = v.ViewerId
WHERE v.DocId = @DocId
ORDER BY v.ViewedAt DESC;";
                vc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var vr = await vc.ExecuteReaderAsync();

                string? GetStringSafe(int ordinal)
                    => vr.IsDBNull(ordinal) ? null : Convert.ToString(vr.GetValue(ordinal));

                while (await vr.ReadAsync())
                {
                    DateTime? viewedUtc = vr.IsDBNull(0) ? (DateTime?)null : vr.GetDateTime(0);
                    DateTime? viewedLocal = viewedUtc?.ToLocalTime();
                    string? when = viewedLocal?.ToString("yyyy-MM-dd HH:mm");

                    viewLogs.Add(new
                    {
                        viewedAt = viewedLocal,
                        viewedAtText = when,
                        userId = GetStringSafe(1),
                        viewerRole = GetStringSafe(2),
                        clientIp = GetStringSafe(3),
                        userAgent = GetStringSafe(4),
                        userName = GetStringSafe(5)
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load view logs failed for {DocId}", id);
            }

            // ---------------- 7) ViewBag 세팅 ----------------
            ViewBag.DocId = id;
            ViewBag.DocumentId = id;
            ViewBag.TemplateCode = templateCode ?? "";
            ViewBag.TemplateTitle = templateTitle ?? "";
            ViewBag.Status = status ?? "";
            ViewBag.CompCd = compCd ?? "";

            ViewBag.CreatorName = creatorDisplayName ?? string.Empty;

            ViewBag.DescriptorJson = descriptorJson ?? "{}";
            ViewBag.InputsJson = inputsJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.ApprovalsJson = JsonSerializer.Serialize(approvals);
            ViewBag.ViewLogsJson = JsonSerializer.Serialize(viewLogs);
            ViewBag.ApprovalCellsJson = JsonSerializer.Serialize(approvalCells);

            ViewBag.DocumentFilesJson = JsonSerializer.Serialize(docFiles);

            ViewBag.SelectedRecipientUserIds = sharedRecipientUserIds;
            ViewBag.SelectedRecipientUserIdsJson = JsonSerializer.Serialize(sharedRecipientUserIds);

            ViewBag.CreatedAt = createdAtForUi;
            ViewBag.CreatedAtText = createdAtText;
            ViewBag.RecalledAt = recalledAtForUi;
            ViewBag.RecalledAtText = recalledAtText;

            var caps = await GetApprovalCapabilitiesAsync(id);
            ViewBag.CanRecall = caps.canRecall;
            ViewBag.CanApprove = caps.canApprove;
            ViewBag.CanHold = caps.canHold;
            ViewBag.CanReject = caps.canReject;
            ViewBag.CanBackToNew = caps.canRecall;

            return View("Detail");
        }

        // (추가) Detail에서 공유자 변경 저장
        [HttpPost("UpdateShares")]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> UpdateShares([FromBody] UpdateSharesDto dto)
        {
            if (dto is null || string.IsNullOrWhiteSpace(dto.DocId))
                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_SaveFailed" }, stage = "arg", detail = "docId null/empty" });

            var docId = dto.DocId!.Trim();
            var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";

            // 선택값 정리(본인 제외, 중복 제거)
            var selected = (dto.SelectedRecipientUserIds ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
                .ToList();

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();
                await using var tx = await conn.BeginTransactionAsync();

                try
                {
                    // 1) 현재 활성 공유자(미회수) 목록 조회
                    var currentActive = new List<string>();
                    await using (var sel = conn.CreateCommand())
                    {
                        sel.Transaction = (SqlTransaction)tx;
                        sel.CommandText = @"
SELECT UserId
FROM dbo.DocumentShares
WHERE DocId = @DocId AND ISNULL(IsRevoked,0) = 0;";
                        sel.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });

                        await using var r = await sel.ExecuteReaderAsync();
                        while (await r.ReadAsync())
                        {
                            var uid = (r["UserId"]?.ToString() ?? "").Trim();
                            if (!string.IsNullOrWhiteSpace(uid))
                                currentActive.Add(uid);
                        }
                    }

                    // 2) 제외된 사람은 revoke 처리 + 로그
                    var toRevoke = currentActive
                        .Where(uid => !selected.Any(x => string.Equals(x, uid, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    foreach (var targetUserId in toRevoke)
                    {
                        await using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = (SqlTransaction)tx;
                            cmd.CommandText = @"
UPDATE dbo.DocumentShares
   SET IsRevoked = 1,
       ExpireAt = SYSUTCDATETIME()
 WHERE DocId = @DocId AND UserId = @UserId;";
                            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
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
                            logCmd.Parameters.Add(new SqlParameter("@ChangeCode", SqlDbType.NVarChar, 50) { Value = "ShareRevoked" });
                            logCmd.Parameters.Add(new SqlParameter("@TargetUserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                            logCmd.Parameters.Add(new SqlParameter("@AfterJson", SqlDbType.NVarChar, -1) { Value = "{\"revoked\":true}" });
                            await logCmd.ExecuteNonQueryAsync();
                        }
                    }

                    // 3) 선택된 사람은 upsert(기존 Create 로직 그대로) + 로그
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

                return Json(new { ok = true, docId = docId, selectedCount = selected.Count });
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "UpdateShares failed docId={docId}", docId);
                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_SaveFailed" }, stage = "db", detail = ex.Message });
            }
        }



        [HttpGet("Download/{fileKey}")]
        public async Task<IActionResult> Download([FromRoute] string fileKey)
        {
            if (string.IsNullOrWhiteSpace(fileKey))
                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT TOP (1) DocId, FileKey, OriginalName, StoragePath, ContentType, ByteSize
FROM dbo.DocumentFiles
WHERE FileKey = @FileKey
ORDER BY UploadedAt DESC, FileKey DESC;";

                cmd.Parameters.Add(new SqlParameter("@FileKey", SqlDbType.NVarChar, 200) { Value = fileKey });

                await using var r = await cmd.ExecuteReaderAsync(CommandBehavior.SingleRow);
                if (!await r.ReadAsync())
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                var storagePath = (r["StoragePath"] as string) ?? string.Empty;
                var originalName = (r["OriginalName"] as string) ?? fileKey;
                var contentType = (r["ContentType"] as string) ?? "application/octet-stream";

                // StoragePath는 반드시 "상대경로"여야 함 (절대경로 금지 정책)
                // 경로 탐색 방지: 루트/드라이브/상위(..) 차단 + ContentRootPath 결합
                var rel = storagePath.Replace('\\', '/').Trim();
                if (string.IsNullOrWhiteSpace(rel))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                if (rel.StartsWith("/") || rel.Contains(":") || rel.Contains(".."))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                var full = Path.GetFullPath(Path.Combine(_env.ContentRootPath, rel));
                var root = Path.GetFullPath(_env.ContentRootPath);

                // ContentRoot 하위만 허용
                if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                if (!System.IO.File.Exists(full))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                // 스트림 반환
                var stream = new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.Read);
                return File(stream, contentType, originalName, enableRangeProcessing: true);
            }
            catch
            {
                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });
            }
        }
        private static string CalculateViewerRole(
            string? createdBy,
            string viewerId,
            List<(int StepOrder, string? UserId, string Status)> approvals)
        {
            if (string.IsNullOrWhiteSpace(viewerId))
                return "Unknown";

            var isCreator = !string.IsNullOrWhiteSpace(createdBy) &&
                            string.Equals(createdBy, viewerId, StringComparison.OrdinalIgnoreCase);

            var isApprover = approvals.Exists(a =>
                !string.IsNullOrWhiteSpace(a.UserId) &&
                string.Equals(a.UserId, viewerId, StringComparison.OrdinalIgnoreCase));

            if (isCreator && isApprover) return "Creator+Approver";
            if (isCreator) return "Creator";
            if (isApprover) return "Approver";
            return "Viewer";
        }

        private async Task LogDocumentViewAsync(SqlConnection conn, string docId)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            var userName = User.Identity?.Name ?? "";
            var clientIp = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "";
            var userAgent = Request.Headers["User-Agent"].ToString() ?? "";

            // ViewerRole 계산(기안자/결재자/공유자 여부 등은 나중에 더 정교하게 바꿔도 됨)
            string viewerRole = "Other";

            // 기안자 / 결재자 여부 간단 판별
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT 
    CreatedBy,
    CASE WHEN EXISTS (
        SELECT 1 
        FROM dbo.DocumentApprovals da 
        WHERE da.DocId = d.DocId AND da.UserId = @UserId
    ) THEN 1 ELSE 0 END AS IsApprover
FROM dbo.Documents d
WHERE d.DocId = @DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = (object)userId ?? DBNull.Value });

                using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    var createdBy = r["CreatedBy"] as string;
                    var isApprover = (int)r["IsApprover"] == 1;

                    var isCreator = !string.IsNullOrWhiteSpace(createdBy) &&
                                    string.Equals(createdBy, userId, StringComparison.OrdinalIgnoreCase);

                    if (isCreator && isApprover)
                        viewerRole = "Creator+Approver";
                    else if (isCreator)
                        viewerRole = "Creator";
                    else if (isApprover)
                        viewerRole = "Approver";
                    else
                        viewerRole = "SharedOrOther";
                }
            }

            // 중복 방지(선택사항): 같은 사용자·역할이 최근 N분 안에 본 기록이 있으면 스킵 등은 나중에 필요하면 추가
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
INSERT INTO dbo.DocumentViewLogs
    (DocId, ViewerId, ViewerRole, ViewedAt, ClientIp, UserAgent)
VALUES
    (@DocId, @ViewerId, @ViewerRole, SYSUTCDATETIME(), @ClientIp, @UserAgent);";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@ViewerId", SqlDbType.NVarChar, 450)
                {
                    Value = userId ?? (object)DBNull.Value
                });
                cmd.Parameters.Add(new SqlParameter("@ViewerRole", SqlDbType.NVarChar, 50) { Value = (object)viewerRole ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@ClientIp", SqlDbType.NVarChar, 64) { Value = (object)clientIp ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@UserAgent", SqlDbType.NVarChar, 512) { Value = (object)userAgent ?? DBNull.Value });

                await cmd.ExecuteNonQueryAsync();
            }

            // 필요하면 여기서 DocumentAuditLogs 에도 "View" 액션으로 1줄 남겨도 됨
        }

        // ---------- Helper: DocumentViewLogs INSERT ----------
        private async Task WriteDocumentViewLogAsync(
            SqlConnection conn,
            string docId,
            string viewerId,
            string viewerRole)
        {
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(viewerId))
                return;

            var ip = HttpContext?.Connection?.RemoteIpAddress?.ToString() ?? string.Empty;
            var ua = HttpContext?.Request?.Headers["User-Agent"].ToString() ?? string.Empty;

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
INSERT INTO dbo.DocumentViewLogs
    (DocId, ViewerId, ViewerRole, ViewedAt, ClientIp, UserAgent)
VALUES
    (@DocId, @ViewerId, @ViewerRole, SYSUTCDATETIME(), @ClientIp, @UserAgent);";

            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@ViewerId", SqlDbType.NVarChar, 450) { Value = viewerId });
            cmd.Parameters.Add(new SqlParameter("@ViewerRole", SqlDbType.NVarChar, 50)
            {
                Value = string.IsNullOrWhiteSpace(viewerRole) ? (object)DBNull.Value : viewerRole
            });
            cmd.Parameters.Add(new SqlParameter("@ClientIp", SqlDbType.NVarChar, 64)
            {
                Value = string.IsNullOrWhiteSpace(ip) ? (object)DBNull.Value : ip
            });
            cmd.Parameters.Add(new SqlParameter("@UserAgent", SqlDbType.NVarChar, 512)
            {
                Value = string.IsNullOrWhiteSpace(ua) ? (object)DBNull.Value : ua
            });

            await cmd.ExecuteNonQueryAsync();
        }

        /// <summary>
        /// Detail 화면 진입 시 DocumentViewLogs + DocumentAuditLogs 에 조회 이력 기록
        /// </summary>
        private async Task InsertDocumentViewAndAuditAsync(
            SqlConnection conn,
            string docId,
            string? userId,
            string? userName)
        {
            if (string.IsNullOrWhiteSpace(docId))
                return;

            // 익명/시스템 계정까지 모두 남길지 여부는 필요 시 조건 추가
            var uid = userId ?? string.Empty;
            var uname = userName ?? string.Empty;

            try
            {
                // 1) DocumentViewLogs: 누가 언제 봤는지
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO dbo.DocumentViewLogs (DocId, ViewedAtUtc, UserId, UserName)
VALUES (@DocId, SYSUTCDATETIME(), @UserId, @UserName);";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450)
                    {
                        Value = string.IsNullOrEmpty(uid) ? DBNull.Value : uid
                    });
                    cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200)
                    {
                        Value = string.IsNullOrEmpty(uname) ? DBNull.Value : uname
                    });

                    await cmd.ExecuteNonQueryAsync();
                }

                // 2) DocumentAuditLogs: 액션 코드만 간단히 남김 (기존 테이블 구조 기준)
                //    - 컬럼: DocId, UserId, ActionCode, CreatedAtUtc(기본값) 정도만 있다고 가정
                await using (var ac = conn.CreateCommand())
                {
                    ac.CommandText = @"
INSERT INTO dbo.DocumentAuditLogs (DocId, UserId, ActionCode)
VALUES (@DocId, @UserId, @ActionCode);";
                    ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    ac.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450)
                    {
                        Value = string.IsNullOrEmpty(uid) ? DBNull.Value : uid
                    });
                    ac.Parameters.Add(new SqlParameter("@ActionCode", SqlDbType.NVarChar, 50) { Value = "View" });

                    await ac.ExecuteNonQueryAsync();
                }
            }
            catch (Exception ex)
            {
                // 조회 로그 기록 실패해도 화면은 계속 보여야 하므로 Warning 으로만 남김
                _log.LogWarning(ex, "Detail: InsertDocumentViewAndAuditAsync failed for {DocId}", docId);
            }
        }

        // ========== DetailsData (상세 데이터 API) ==========
        [HttpGet("DetailsData")]
        [Produces("application/json")]
        public async Task<IActionResult> DetailsData(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return BadRequest(new { message = "DOC_Err_BadRequest" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // ----- 1) 문서 기본 정보 -----
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT d.DocId,
       d.TemplateCode,
       d.TemplateTitle,
       d.Status,
       d.DescriptorJson,
       d.OutputPath,
       d.CreatedAt
FROM dbo.Documents d
WHERE d.DocId = @id;";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

            string? descriptor = null, output = null, title = null, code = null, status = null;
            DateTime? createdAt = null;

            await using (var rd = await cmd.ExecuteReaderAsync())
            {
                if (await rd.ReadAsync())
                {
                    code = rd["TemplateCode"] as string ?? "";
                    title = rd["TemplateTitle"] as string ?? "";
                    status = rd["Status"] as string ?? "";
                    descriptor = rd["DescriptorJson"] as string ?? "{}";
                    output = rd["OutputPath"] as string ?? "";
                    if (!rd.IsDBNull(rd.GetOrdinal("CreatedAt")))
                        createdAt = (DateTime)rd["CreatedAt"];
                }
                else
                {
                    return NotFound(new { message = "DOC_Err_DocumentNotFound" });
                }
            }

            var preview = string.IsNullOrWhiteSpace(output) || !System.IO.File.Exists(output)
                ? "{}"
                : BuildPreviewJsonFromExcel(output);

            // ----- 2) 결재 정보 DisplayName 조인 -----
            var approvals = new List<object>();
            await using (var ac = conn.CreateCommand())
            {
                ac.CommandText = @"
SELECT  a.StepOrder,
        a.RoleKey,
        a.ApproverValue,
        a.UserId,
        a.Status,
        a.Action,
        a.ActedAt,
        a.ActorName,
        up.DisplayName
FROM dbo.DocumentApprovals a
LEFT JOIN dbo.UserProfiles up
       ON up.UserId = a.UserId
WHERE a.DocId = @DocId
ORDER BY a.StepOrder;";
                ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    DateTime? acted = r.IsDBNull(6) ? (DateTime?)null : r.GetDateTime(6);
                    string? when = acted.HasValue
                        ? ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                        : null;

                    var approverValue = r["ApproverValue"]?.ToString();
                    var userId = r["UserId"]?.ToString();
                    var actorName = r["ActorName"]?.ToString();
                    var displayName = r["DisplayName"] as string ?? string.Empty;

                    // UserProfiles DisplayName 이 있으면 그 값을 표시용 이름으로 사용
                    var approverDisplayText =
                        !string.IsNullOrWhiteSpace(displayName) ? displayName :
                        !string.IsNullOrWhiteSpace(actorName) ? actorName :
                        approverValue ?? string.Empty;

                    approvals.Add(new
                    {
                        step = r.GetInt32(0),
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue,
                        userId,
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = when,
                        actorName,
                        ApproverDisplayText = approverDisplayText
                    });
                }
            }

            // ----- 3) 결과 반환 -----
            return Json(new
            {
                ok = true,
                docId = id,
                templateCode = code,
                templateTitle = title,
                status,
                createdAt = createdAt.HasValue
                    ? ToLocalStringFromUtc(DateTime.SpecifyKind(createdAt.Value, DateTimeKind.Utc))
                    : null,
                descriptorJson = descriptor,
                previewJson = preview,
                approvals
            });
        }
        private string ResolveTempUploadPath(string fileKey)
        {
            if (string.IsNullOrWhiteSpace(fileKey))
                return string.Empty;

            // path traversal 방지: 파일명만 사용
            var safeKey = Path.GetFileName(fileKey.Trim());

            return Path.Combine(_env.ContentRootPath, "App_Data", "Uploads", "Temp", safeKey);
        }

        // ========= DTOs =========
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

        // (추가) Detail 공유자 변경용 DTO
        public sealed class UpdateSharesDto
        {
            public string? DocId { get; set; }
            public List<string>? SelectedRecipientUserIds { get; set; }
        }

        // ========= Claims → 회사/부서 =========
        private (string compCd, string? departmentId) GetUserCompDept()
        {
            var comp = User.FindFirstValue("compCd");
            var dept = User.FindFirstValue("departmentId");

            if (!string.IsNullOrWhiteSpace(comp))
                return (comp, string.IsNullOrWhiteSpace(dept) ? null : dept);

            try
            {
                var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
                if (!string.IsNullOrEmpty(userId))
                {
                    var cs = _cfg.GetConnectionString("DefaultConnection");
                    using var conn = new SqlConnection(cs);
                    conn.Open();
                    using var cmd = new SqlCommand(
                        @"SELECT TOP 1 CompCd, DepartmentId FROM dbo.UserProfiles WHERE UserId=@uid", conn);
                    cmd.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 450) { Value = userId });
                    using var rd = cmd.ExecuteReader();
                    if (rd.Read())
                    {
                        var c = rd.IsDBNull(0) ? null : rd.GetString(0);
                        var d = rd.IsDBNull(1) ? null : rd.GetInt32(1).ToString();
                        return (c ?? "", d);
                    }
                }
            }
            catch { }
            return ("", string.IsNullOrWhiteSpace(dept) ? null : dept);
        }

        // ========= Preview 생성 =========
        private static string BuildPreviewJsonFromExcel(string excelPath, int maxRows = 50, int maxCols = 26)
        {
            using var wb = new XLWorkbook(excelPath);
            var ws0 = wb.Worksheets.First();

            double defaultRowPt = ws0.RowHeight; if (defaultRowPt <= 0) defaultRowPt = 15.0;
            double defaultColChar = ws0.ColumnWidth; if (defaultColChar <= 0) defaultColChar = 8.43;

            var cells = new List<List<string>>(maxRows);
            for (int r = 1; r <= maxRows; r++)
            {
                var row = new List<string>(maxCols);
                for (int c = 1; c <= maxCols; c++)
                {
                    var cell = ws0.Cell(r, c);
                    row.Add(cell.GetFormattedString());
                }
                cells.Add(row);
            }

            var merges = new List<int[]>();
            foreach (var mr in ws0.MergedRanges)
            {
                var a = mr.RangeAddress;
                int r1 = a.FirstAddress.RowNumber, c1 = a.FirstAddress.ColumnNumber;
                int r2 = a.LastAddress.RowNumber, c2 = a.LastAddress.ColumnNumber;
                if (r1 > maxRows || c1 > maxCols) continue;
                r2 = Math.Min(r2, maxRows);
                c2 = Math.Min(c2, maxCols);
                if (r1 <= r2 && c1 <= c2) merges.Add(new[] { r1, c1, r2, c2 });
            }

            var colW = new List<double>(maxCols);
            for (int c = 1; c <= maxCols; c++)
            {
                var w = ws0.Column(c).Width;
                if (w <= 0) w = defaultColChar;
                colW.Add(w);
            }

            var rowH = new List<double>(maxRows);
            for (int r = 1; r <= maxRows; r++)
            {
                var h = ws0.Row(r).Height;
                if (h <= 0) h = defaultRowPt;
                rowH.Add(h);
            }

            var styles = new Dictionary<string, object>();
            for (int r = 1; r <= maxRows; r++)
            {
                for (int c = 1; c <= maxCols; c++)
                {
                    var cell = ws0.Cell(r, c);
                    var st = cell.Style;
                    string? bgHex = ToHexIfRgb(cell);
                    styles[$"{r},{c}"] = new
                    {
                        font = new
                        {
                            name = st.Font.FontName,
                            size = st.Font.FontSize,
                            bold = st.Font.Bold,
                            italic = st.Font.Italic,
                            underline = st.Font.Underline != XLFontUnderlineValues.None
                        },
                        align = new
                        {
                            h = st.Alignment.Horizontal.ToString(),
                            v = st.Alignment.Vertical.ToString(),
                            wrap = st.Alignment.WrapText
                        },
                        border = new
                        {
                            l = st.Border.LeftBorder.ToString(),
                            r = st.Border.RightBorder.ToString(),
                            t = st.Border.TopBorder.ToString(),
                            b = st.Border.BottomBorder.ToString()
                        },
                        fill = new { bg = bgHex }
                    };
                }
            }

            return JsonSerializer.Serialize(new
            {
                sheet = ws0.Name,
                rows = maxRows,
                cols = maxCols,
                cells,
                merges,
                colW,
                rowH,
                styles
            });
        }

        private static string? ToHexIfRgb(IXLCell cell)
        {
            try
            {
                var bg = cell?.Style?.Fill?.BackgroundColor;
                if (bg != null && bg.ColorType == XLColorType.Color)
                {
                    var c = bg.Color;
                    return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                }
            }
            catch { }
            return null;
        }

        private static bool TryParseJsonFlexible(string? json, out JsonDocument doc)
        {
            doc = null!;
            if (string.IsNullOrWhiteSpace(json)) return false;

            try
            {
                var first = JsonDocument.Parse(json);
                if (first.RootElement.ValueKind == JsonValueKind.String)
                {
                    var inner = first.RootElement.GetString();
                    first.Dispose();
                    if (string.IsNullOrWhiteSpace(inner)) return false;
                    doc = JsonDocument.Parse(inner!);
                    return true;
                }
                doc = first;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// 현재 로그인 사용자가 주어진 문서에 대해  회수 가능 여부(CanRecall) 승인/보류/반려 가능 여부(CanApprove)를 판단.
        // 2025.11.20 Changed: GetApprovalCapabilitiesAsync 에 Hold/Reject 권한 플래그 추가 (튜플 2개 → 4개로 확장, Detail 뷰와 정합)
        
        // ========= 입력값 → 엑셀 =========
        private async Task<string> GenerateExcelFromInputsAsync(
            long versionId,
            string templateExcelPath,
            Dictionary<string, string> inputs,
            string? descriptorVersion)
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

        private static (string BaseKey, int? Index) ParseKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return ("", null);
            var i = key.LastIndexOf('_');
            if (i > 0 && int.TryParse(key[(i + 1)..], out var n))
                return (key[..i], n);
            return (key, null);
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

        // ========= 수신자 해석 =========
        private (List<string> emails, List<string> diag) GetInitialRecipients(
            ComposePostDto dto, string normalizedDesc, IConfiguration cfg, string? docId = null, string? templateCode = null)
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

        // ========= 저장 =========
        // 2025.12.08 Changed: 템플릿 기반 결재 슬롯 생성 로직 제거하고, 이전처럼 approvalsJson 배열만 사용하도록 롤백
        // 2025.12.08 Changed: 문서 저장 시 DocTemplateApproval 슬롯 기준으로 누락된 결재단계를 자동 보정하여 DocumentApprovals 총 개수가 템플릿 슬롯 수와 일치하도록 수정
        private async Task SaveDocumentAsync(
    string docId,
    string templateCode,
    string title,
    string status,
    string outputPath,
    Dictionary<string, string> inputs,
    string approvalsJson,
    string descriptorJson,
    string userId,
    string userName,
    string compCd,
    string? departmentId,
    int? templateVersionId = null
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

        private static async Task<int?> ResolveTemplateVersionIdAsync(
            SqlConnection conn,
            SqlTransaction tx,
            string templateCode
        )
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

                            // RoleKey : roleKey / RoleKey / part / lineType 등을 우선순위로 사용
                            string roleKey = string.Empty;

                            if (a.TryGetProperty("roleKey", out var rk1) && rk1.ValueKind == JsonValueKind.String)
                                roleKey = rk1.GetString() ?? string.Empty;
                            else if (a.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String)
                                roleKey = rk2.GetString() ?? string.Empty;
                            else if (a.TryGetProperty("part", out var rk3) && rk3.ValueKind == JsonValueKind.String)
                                roleKey = rk3.GetString() ?? string.Empty;

                            if (string.IsNullOrWhiteSpace(roleKey))
                                roleKey = $"A{stepOrder}";

                            // ApproverValue : value / approverValue / ApproverValue 중 첫 번째 문자열
                            string? approverValue = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String)
                                approverValue = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String)
                                approverValue = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String)
                                approverValue = v3.GetString();

                            approvals.Add((stepOrder, roleKey, approverValue));
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

                // 3) 기존 DocumentApprovals 모두 삭제(신규 생성 시에는 원래 0건이지만 안전하게 정리)
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

                foreach (var a in ordered)
                {
                    await using var ins = conn.CreateCommand();
                    ins.Transaction = (SqlTransaction)tx;
                    ins.CommandText = @"
INSERT INTO dbo.DocumentApprovals
    (DocId, StepOrder, RoleKey, ApproverValue, Status, CreatedAt)
VALUES
    (@DocId, @StepOrder, @RoleKey, @ApproverValue, N'Pending', SYSUTCDATETIME());";

                    ins.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    ins.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = a.StepOrder });
                    ins.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 100) { Value = a.RoleKey });

                    var pVal = new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128);
                    pVal.Value = string.IsNullOrWhiteSpace(a.ApproverValue)
                        ? (object)DBNull.Value
                        : a.ApproverValue!;
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
                try { await tx.RollbackAsync(); } catch { /* ignore */ }
                _log.LogError(ex, "EnsureApprovalsAndSyncAsync failed docId={docId}, docCode={docCode}", docId, docCode);
            }
        }

//        // ========= Board =========
//        [HttpGet("Board")]
//        public IActionResult Board()
//        {
//            return View("Board");
//        }

//        [HttpGet("BoardData")]
//        // 2025.12.23 Changed: 공유 문서함 조회에 DocumentShares IsRevoked 조건과 만료 조건을 추가하여 철회 및 만료 공유를 제외
//        // 2025.12.09 Changed: 결재 문서함 승인 탭에서 선행 단계 보류 반려 문서가 후행 승인자에게 표시되지 않도록 필터 보강
//        public async Task<IActionResult> BoardData(
//    string tab = "created",
//    int page = 1,
//    int pageSize = 20,
//    string titleFilter = "all",
//    string sort = "created_desc",
//    string? q = null,
//    string approvalView = "all")
//        {
//            if (page < 1) page = 1;
//            if (pageSize < 1 || pageSize > 100) pageSize = 20;

//            var userId = User?.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;

//            string orderBy = sort switch
//            {
//                "created_asc" => "ORDER BY d.CreatedAt ASC",
//                "title_asc" => "ORDER BY d.TemplateTitle ASC",
//                "title_desc" => "ORDER BY d.TemplateTitle DESC",
//                _ => "ORDER BY d.CreatedAt DESC"
//            };

//            string whereSearch = string.IsNullOrWhiteSpace(q) ? "" : " AND d.TemplateTitle LIKE @Q ";

//            // 2025.12.08 Changed: 보류/반려 문서가 OnHoldA1 / RejectedA1 처럼 저장된 경우도 필터에 포함되도록 LIKE 로 변경
//            string whereTitleFilter = titleFilter?.ToLowerInvariant() switch
//            {
//                "approved" => " AND ISNULL(d.Status, N'') LIKE N'Approved%' ",
//                "rejected" => " AND ISNULL(d.Status, N'') LIKE N'Rejected%' ",
//                "pending" => " AND ISNULL(d.Status, N'') LIKE N'Pending%' ",
//                "onhold" => " AND ISNULL(d.Status, N'') LIKE N'OnHold%' ",
//                "recalled" => " AND ISNULL(d.Status, N'') LIKE N'Recalled%' ",
//                _ => ""
//            };

//            string sqlCount;
//            string sqlList;
//            var offset = (page - 1) * pageSize;

//            if (tab == "created")
//            {
//                sqlCount = @"
//SELECT COUNT(1)
//FROM Documents d
//WHERE d.CreatedBy = @UserId" + whereSearch + whereTitleFilter + ";";

//                sqlList = @"
//SELECT d.DocId,
//       d.TemplateTitle,
//       d.CreatedAt,
//       ISNULL(d.Status, N'') AS Status,
//       up.DisplayName AS AuthorName,
//       (SELECT COUNT(1)
//          FROM DocumentComments c
//         WHERE c.DocId = d.DocId
//           AND c.IsDeleted = 0) AS CommentCount,
//       CAST(0 AS bit) AS HasAttachment,
//       (SELECT COUNT(1)
//          FROM DocumentApprovals da
//         WHERE da.DocId = d.DocId) AS TotalSteps,
//       (SELECT COUNT(1)
//          FROM DocumentApprovals da
//         WHERE da.DocId = d.DocId
//           AND (ISNULL(da.Action, N'') = N'Approved'
//             OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps
//FROM Documents d
//LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
//WHERE d.CreatedBy = @UserId" + whereSearch + whereTitleFilter + $@"
//{orderBy}
//OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
//            }
//            else if (tab == "approval")
//            {
//                // 2025.12.08 기존 코드 유지
//                string whereApprovalView = (approvalView ?? string.Empty).ToLowerInvariant() switch
//                {
//                    "pending" =>
//                        " AND ISNULL(a.Status, N'') = N'Pending' " +
//                        " AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10)) ",
//                    "approved" =>
//                        " AND (ISNULL(a.Action, N'') = N'Approved' OR ISNULL(a.Status, N'') = N'Approved') ",
//                    _ => @"
//  AND (
//        (
//          ISNULL(a.Status, N'') = N'Pending'
//          AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10))
//        )
//        OR
//        (
//          ISNULL(a.Action, N'') <> N''
//          OR (
//               ISNULL(a.Status, N'') <> N''
//           AND ISNULL(a.Status, N'') <> N'Pending'
//             )
//        )
//      )"
//                };

//                const string whereRecalledFilter = @"
//  AND ISNULL(d.Status, N'') <> N'Recalled'
//  AND ISNULL(a.Status, N'') <> N'Recalled'";

//                sqlCount = @"
//SELECT COUNT(1)
//FROM DocumentApprovals a
//JOIN Documents d ON a.DocId = d.DocId
//WHERE a.UserId = @UserId" + whereRecalledFilter + whereApprovalView + whereSearch + whereTitleFilter + @";";

//                sqlList = @"
//SELECT d.DocId,
//       d.TemplateTitle,
//       d.CreatedAt,
//       CASE 
//           WHEN ISNULL(a.Action, N'') <> N'' THEN a.Action 
//           ELSE ISNULL(d.Status, N'') 
//       END AS Status,
//       up.DisplayName AS AuthorName,
//       (SELECT COUNT(1)
//          FROM DocumentComments c
//         WHERE c.DocId = d.DocId
//           AND c.IsDeleted = 0) AS CommentCount,
//       CAST(0 AS bit) AS HasAttachment,
//       (SELECT COUNT(1)
//          FROM DocumentApprovals da
//         WHERE da.DocId = d.DocId) AS TotalSteps,
//       (SELECT COUNT(1)
//          FROM DocumentApprovals da
//         WHERE da.DocId = d.DocId
//           AND (ISNULL(da.Action, N'') = N'Approved'
//             OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps
//FROM DocumentApprovals a
//JOIN Documents d ON a.DocId = d.DocId
//LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
//WHERE a.UserId = @UserId" + whereRecalledFilter + whereApprovalView + whereSearch + whereTitleFilter + $@"
//{orderBy}
//OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
//            }
//            else // shared
//            {
//                // 2025.12.23 Changed: 공유 문서함은 철회(IsRevoked=0) + 만료(ExpireAt) 조건을 반드시 포함
//                // - 회사/부서/상태 스코프 필터는 shared 에 과도하게 적용하지 않도록 분리(현재 로직은 상태/검색만 적용)
//                const string whereShareActive = @"
// AND ISNULL(s.IsRevoked, 0) = 0
// AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME())";

//                sqlCount = @"
//SELECT COUNT(1)
//FROM DocumentShares s
//JOIN Documents d ON s.DocId = d.DocId
//WHERE s.UserId = @UserId" + whereShareActive + whereSearch + whereTitleFilter + ";";

//                sqlList = @"
//SELECT d.DocId,
//       d.TemplateTitle,
//       d.CreatedAt,
//       ISNULL(d.Status, N'') AS Status,
//       up.DisplayName AS AuthorName,
//       (SELECT COUNT(1)
//          FROM DocumentComments c
//         WHERE c.DocId = d.DocId
//           AND c.IsDeleted = 0) AS CommentCount,
//       CAST(0 AS bit) AS HasAttachment,
//       (SELECT COUNT(1)
//          FROM DocumentApprovals da
//         WHERE da.DocId = d.DocId) AS TotalSteps,
//       (SELECT COUNT(1)
//          FROM DocumentApprovals da
//         WHERE da.DocId = d.DocId
//           AND (ISNULL(da.Action, N'') = N'Approved'
//             OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps
//FROM DocumentShares s
//JOIN Documents d ON s.DocId = d.DocId
//LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
//WHERE s.UserId = @UserId" + whereShareActive + whereSearch + whereTitleFilter + $@"
//{orderBy}
//OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
//            }

//            var connStr = _cfg.GetConnectionString("DefaultConnection");
//            using var conn = new SqlConnection(connStr);
//            await conn.OpenAsync();

//            using var cmdCount = new SqlCommand(sqlCount, conn);
//            cmdCount.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
//            if (!string.IsNullOrWhiteSpace(q))
//                cmdCount.Parameters.Add(new SqlParameter("@Q", SqlDbType.NVarChar, 200) { Value = $"%{q}%" });

//            var total = Convert.ToInt32(await cmdCount.ExecuteScalarAsync());

//            using var cmdList = new SqlCommand(sqlList, conn);
//            cmdList.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
//            if (!string.IsNullOrWhiteSpace(q))
//                cmdList.Parameters.Add(new SqlParameter("@Q", SqlDbType.NVarChar, 200) { Value = $"%{q}%" });
//            cmdList.Parameters.Add(new SqlParameter("@Offset", SqlDbType.Int) { Value = offset });
//            cmdList.Parameters.Add(new SqlParameter("@PageSize", SqlDbType.Int) { Value = pageSize });

//            var items = new List<object>();
//            using (var rdr = await cmdList.ExecuteReaderAsync())
//            {
//                while (await rdr.ReadAsync())
//                {
//                    var createdAtLocal = (rdr["CreatedAt"] is DateTime utc)
//                        ? ToLocalStringFromUtc(utc)
//                        : string.Empty;

//                    var rawStatus = rdr["Status"]?.ToString() ?? string.Empty;

//                    var totalSteps = rdr["TotalSteps"] is int ts ? ts : Convert.ToInt32(rdr["TotalSteps"]);
//                    var completedSteps = rdr["CompletedSteps"] is int cs ? cs : Convert.ToInt32(rdr["CompletedSteps"]);
//                    var commentCount = rdr["CommentCount"] is int cc ? cc : Convert.ToInt32(rdr["CommentCount"]);

//                    items.Add(new
//                    {
//                        docId = rdr["DocId"]?.ToString() ?? string.Empty,
//                        templateTitle = rdr["TemplateTitle"]?.ToString() ?? string.Empty,
//                        authorName = rdr["AuthorName"]?.ToString() ?? string.Empty,
//                        createdAt = createdAtLocal,
//                        status = rawStatus,
//                        statusCode = rawStatus,
//                        totalApprovers = totalSteps,
//                        completedApprovers = completedSteps,
//                        commentCount = commentCount,
//                        hasAttachment = Convert.ToInt32(rdr["HasAttachment"]) == 1
//                    });
//                }
//            }

//            return Json(new { total, page, pageSize, items });
//        }



//        [HttpGet("BoardBadges")]
//        public async Task<IActionResult> BoardBadges()
//        {
//            var userId = User?.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
//            var connStr = _cfg.GetConnectionString("DefaultConnection");
//            using var conn = new SqlConnection(connStr);
//            await conn.OpenAsync();

//            // 2025.12.08 Changed: 회수된 문서는 결재 대기 배지에서 제외되도록 Documents 와 조인
//            // 2025.12.09 Changed: a.Status=Pending 인 행 중에서 실제 현재 차수(PendingA{StepOrder})에 해당하는 것만 집계
//            var sqlApprovalPending = @"
//SELECT COUNT(1)
//FROM DocumentApprovals a
//JOIN Documents d ON a.DocId = d.DocId
//WHERE a.UserId = @UserId
//  AND ISNULL(a.Status, N'') = N'Pending'
//  AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10))
//  AND ISNULL(d.Status, N'') <> N'Recalled';";

//            var sqlSharedUnread = @"
//SELECT COUNT(1)
//FROM DocumentShares s
//WHERE s.UserId = @UserId
//AND NOT EXISTS (
//    SELECT 1 
//    FROM DocumentAuditLogs l
//    WHERE l.DocId = s.DocId AND l.ActorId = @UserId AND l.ActionCode = N'READ'
//);";

//            int approvalPending = 0;
//            int sharedUnread = 0;

//            using (var cmd = new SqlCommand(sqlApprovalPending, conn))
//            {
//                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
//                approvalPending = Convert.ToInt32(await cmd.ExecuteScalarAsync());
//            }
//            using (var cmd = new SqlCommand(sqlSharedUnread, conn))
//            {
//                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
//                sharedUnread = Convert.ToInt32(await cmd.ExecuteScalarAsync());
//            }

//            return Json(new { created = 0, approvalPending, sharedUnread });
//        }

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

        // ========= 시간대/유틸 =========
        private string GetCurrentUserEmail()
        {
            try
            {
                var uid = User?.FindFirstValue(ClaimTypes.NameIdentifier);
                if (string.IsNullOrWhiteSpace(uid)) return string.Empty;

                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();
                using var cmd = new SqlCommand("SELECT TOP 1 Email FROM dbo.AspNetUsers WHERE Id=@id", conn);
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 450) { Value = uid });
                return cmd.ExecuteScalar() as string ?? string.Empty;
            }
            catch { return string.Empty; }
        }

        private string GetCurrentUserDisplayNameStrict()
        {
            try
            {
                var uid = User?.FindFirstValue(ClaimTypes.NameIdentifier);
                if (string.IsNullOrWhiteSpace(uid)) return string.Empty;

                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand(@"
SELECT TOP 1 LTRIM(RTRIM(p.DisplayName))
FROM dbo.UserProfiles p
WHERE p.UserId = @uid;", conn);
                cmd.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 450) { Value = uid });

                var name = cmd.ExecuteScalar() as string;
                return string.IsNullOrWhiteSpace(name) ? string.Empty : name!;
            }
            catch { return string.Empty; }
        }

        private string GetDisplayNameByEmailStrict(string? email)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;

            try
            {
                var em = email.Trim();
                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand(@"
SELECT TOP 1 LTRIM(RTRIM(p.DisplayName))
FROM dbo.AspNetUsers u
JOIN dbo.UserProfiles p ON p.UserId = u.Id
WHERE LTRIM(RTRIM(u.Email)) = @em OR u.NormalizedEmail = UPPER(@em);", conn);
                cmd.Parameters.Add(new SqlParameter("@em", SqlDbType.NVarChar, 256) { Value = em });

                var name = cmd.ExecuteScalar() as string;
                return string.IsNullOrWhiteSpace(name) ? string.Empty : name!;
            }
            catch { return string.Empty; }
        }

        private (string subject, string bodyHtml) BuildSubmissionMail(string authorName, string docTitle, string docId, string? recipientDisplay)
        {
            // 기존 계산 코드는 모두 생략하고, 공통 오버로드로 위임
            return BuildSubmissionMail(authorName, docTitle, docId, recipientDisplay, null);
        }

        private (string subject, string bodyHtml) BuildSubmissionMail(
            string authorName,
            string docTitle,
            string docId,
            string? recipientDisplay,
            DateTime? createdAtLocalForMail = null)
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

        private TimeZoneInfo ResolveCompanyTimeZone()
        {
            var compCd = User.FindFirstValue("compCd") ?? "";

            string? tzIdFromDb = null;
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection");
                if (!string.IsNullOrWhiteSpace(cs) && !string.IsNullOrWhiteSpace(compCd))
                {
                    using var conn = new SqlConnection(cs);
                    conn.Open();
                    using var cmd = new SqlCommand(
                        @"SELECT TOP 1 TimeZoneId FROM dbo.CompMasters WHERE CompCd = @comp", conn);
                    cmd.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar, 10) { Value = compCd });
                    tzIdFromDb = cmd.ExecuteScalar() as string;
                }
            }
            catch { }

            var candidates = new List<string>();
            if (!string.IsNullOrWhiteSpace(tzIdFromDb)) candidates.Add(tzIdFromDb.Trim());
            candidates.Add("Asia/Seoul");
            candidates.Add("Korea Standard Time");

            foreach (var id in candidates)
            {
                try { return TimeZoneInfo.FindSystemTimeZoneById(id); }
                catch
                {
                    var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["Asia/Seoul"] = "Korea Standard Time",
                        ["Korea Standard Time"] = "Asia/Seoul",
                        ["Asia/Ho_Chi_Minh"] = "SE Asia Standard Time",
                        ["Asia/Jakarta"] = "SE Asia Standard Time",
                    };
                    if (map.TryGetValue(id, out var mapped))
                    {
                        try { return TimeZoneInfo.FindSystemTimeZoneById(mapped); } catch { }
                    }
                }
            }

            try { return TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time"); } catch { }
            try { return TimeZoneInfo.FindSystemTimeZoneById("Asia/Seoul"); } catch { }
            return TimeZoneInfo.Utc;
        }

        private string ResolveTimeZoneIdForCurrentUser()
        {
            var tzFromClaim = User?.Claims?.FirstOrDefault(c => c.Type == "TimeZoneId")?.Value;
            if (!string.IsNullOrWhiteSpace(tzFromClaim)) return tzFromClaim;
            return "Korea Standard Time";
        }

        private string ToLocalStringFromUtc(DateTime utc)
        {
            if (utc.Kind != DateTimeKind.Utc)
                utc = DateTime.SpecifyKind(utc, DateTimeKind.Utc);

            var tzId = ResolveTimeZoneIdForCurrentUser();
            TimeZoneInfo tzi;
            try
            {
                tzi = TimeZoneInfo.FindSystemTimeZoneById(tzId);
            }
            catch
            {
                tzi = TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time");
            }

            var local = TimeZoneInfo.ConvertTimeFromUtc(utc, tzi);
            return local.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }

        private static string FallbackNameFromEmail(string? email)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;
            var at = email.IndexOf('@');
            return at > 0 ? email[..at] : email;
        }

        private string ComposeAddress(string? email, string? displayName)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;
            var name = (displayName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(name)) return email.Trim();
            if (email.Contains('<') && email.Contains('>')) return email.Trim();
            return $"{name} <{email.Trim()}>";
        }




        // 2025.12.09 Changed: 문서 생성 시 2차 이상 승인자 이메일을 DocumentApprovals에 매핑하고 상세 디버그 로그 추가

        // (중략) SaveDocumentAsync 안에서 호출 부분은 그대로 두시고,
        // EnsureApprovalsAndSyncAsync 바로 다음에 이 메서드를 호출하고 있는 현재 위치를 유지하시면 됩니다.
        //   await EnsureApprovalsAndSyncAsync(docId, tc);
        //   await FillDocumentApprovalsFromEmailsAsync(docId, toEmails);

        /// <summary>
        /// 문서 상신 시 메일 발송 대상 목록(toEmails)을 이용해
        /// DocumentApprovals.ApproverValue / UserId 를 StepOrder 순서대로 채워 넣는 디버그용 메서드
        /// </summary>
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
        private async Task<(string docId, string outputPath)> SaveDocumentWithCollisionGuardAsync(
            string docId,
            string templateCode,
            string title,
            string status,
            string outputPath,
            Dictionary<string, string> inputs,
            string approvalsJson,
            string descriptorJson,
            string userId,
            string userName,
            string compCd,
            string? departmentId)
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

        [HttpGet("Comments")]
        [Produces("application/json")]
        public async Task<IActionResult> GetComments(
            [FromQuery] string? docId,
            [FromQuery] string? id,
            [FromQuery] string? documentId)
        {
            docId = !string.IsNullOrWhiteSpace(docId)
                ? docId
                : !string.IsNullOrWhiteSpace(id)
                    ? id
                    : documentId;

            if (string.IsNullOrWhiteSpace(docId))
                return Json(new { items = Array.Empty<object>() });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var items = new List<object>();

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            var childMap = new Dictionary<long, int>();
            {
                using var childCmd = conn.CreateCommand();
                childCmd.CommandText = @"
SELECT ParentCommentId, COUNT(*) AS ChildCount
FROM dbo.DocumentComments
WHERE DocId = @DocId AND IsDeleted = 0 AND ParentCommentId IS NOT NULL
GROUP BY ParentCommentId;";
                childCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                using var childRd = await childCmd.ExecuteReaderAsync();
                while (await childRd.ReadAsync())
                {
                    var pid = Convert.ToInt64(childRd["ParentCommentId"]);
                    var cnt = Convert.ToInt32(childRd["ChildCount"]);
                    childMap[pid] = cnt;
                }
            }

            var cmd = conn.CreateCommand();

            cmd.CommandText = @"
SELECT  c.CommentId,
        c.DocId,
        c.ParentCommentId,
        c.ThreadRootId,
        c.Depth,
        c.Body,
        c.HasAttachment,
        c.IsDeleted,
        c.CreatedBy,
        c.CreatedAt,
        c.UpdatedAt,
        COALESCE(p.DisplayName, u.UserName, u.Email, c.CreatedBy) AS AuthorName
FROM dbo.DocumentComments c
LEFT JOIN dbo.AspNetUsers  u
       ON  u.Id       = c.CreatedBy
       OR  u.UserName = c.CreatedBy
       OR  u.Email    = c.CreatedBy
LEFT JOIN dbo.UserProfiles p
       ON  p.UserId   = u.Id
WHERE c.DocId = @DocId
  AND c.IsDeleted = 0
ORDER BY c.ThreadRootId, c.Depth, c.CreatedAt;";

            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

            var currentUserName = User?.Identity?.Name ?? "";
            var currentUserId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            var isAdmin = User?.IsInRole("Admin") ?? false;

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var commentId = Convert.ToInt64(rd["CommentId"]);
                var parent = rd["ParentCommentId"] == DBNull.Value
                    ? (long?)null
                    : Convert.ToInt64(rd["ParentCommentId"]);
                var depth = Convert.ToInt32(rd["Depth"]);
                var hasAttachment = Convert.ToBoolean(rd["HasAttachment"]);

                var createdAtUtc = (DateTime)rd["CreatedAt"];
                var createdAtLocal = ToLocalStringFromUtc(
                    DateTime.SpecifyKind(createdAtUtc, DateTimeKind.Utc));

                var createdByRaw = rd["CreatedBy"]?.ToString() ?? "";

                // 2025.12.08 Changed: 예전 데이터(사용자 Id 저장)와 신규 데이터(UserName 저장)를 모두 허용
                var canModify =
                    (!string.IsNullOrWhiteSpace(createdByRaw) &&
                     (string.Equals(createdByRaw, currentUserName, StringComparison.OrdinalIgnoreCase) ||
                      string.Equals(createdByRaw, currentUserId, StringComparison.OrdinalIgnoreCase)))
                    || isAdmin;

                var authorName = rd["AuthorName"]?.ToString() ?? createdByRaw;
                var hasChildren = childMap.ContainsKey(commentId);

                items.Add(new
                {
                    commentId,
                    docId = (string)rd["DocId"],
                    parentCommentId = parent,
                    depth,
                    body = rd["Body"]?.ToString() ?? "",
                    hasAttachment,
                    createdBy = authorName,
                    createdAt = createdAtLocal,
                    canEdit = canModify && !hasChildren,
                    canDelete = canModify && !hasChildren,
                    hasChildren
                });
            }

            return Json(new { items });
        }

        [HttpPost("Comments")]
        [ValidateAntiForgeryToken] // JSON 요청도 EB-CSRF 헤더로 보호
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> PostComment([FromBody] DocCommentPostDto dto)
        {
            try
            {
                if (dto == null || string.IsNullOrWhiteSpace(dto.docId))
                    return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" }, detail = "docId missing" });

                var body = (dto.body ?? "").Trim();
                if (string.IsNullOrWhiteSpace(body))
                    return BadRequest(new { messages = new[] { "DOC_Val_Required" }, detail = "body empty" });

                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
                var userName = User?.Identity?.Name ?? userId;

                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                int? parentId = dto.parentCommentId;
                int threadRootId;
                int depth;

                if (parentId.HasValue && parentId.Value > 0)
                {
                    var q = conn.CreateCommand();
                    q.CommandText = @"
SELECT ISNULL(ThreadRootId, CommentId) AS RootId, Depth
FROM dbo.DocumentComments
WHERE CommentId = @pid AND DocId = @docId;";
                    q.Parameters.Add(new SqlParameter("@pid", SqlDbType.Int) { Value = parentId.Value });
                    q.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = dto.docId });

                    await using var r = await q.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        // 2025.12.04 Changed: bigint tinyint 을 object 에서 바로 (int) 캐스팅하던 부분을 Convert 사용으로 변경 InvalidCastException 방지
                        threadRootId = Convert.ToInt32(r["RootId"]);
                        depth = Convert.ToInt32(r["Depth"]) + 1;
                    }
                    else
                    {
                        threadRootId = parentId.Value;
                        depth = 1;
                    }
                }
                else
                {
                    threadRootId = 0;
                    depth = 0;
                    parentId = null;
                }

                bool hasAttachment = dto.files != null && dto.files.Count > 0;

                var cmd = conn.CreateCommand();
                cmd.CommandText = @"
INSERT INTO dbo.DocumentComments
    (DocId, ParentCommentId, ThreadRootId, Depth,
     Body, HasAttachment, IsDeleted,
     CreatedBy, CreatedAt, UpdatedAt)
OUTPUT INSERTED.CommentId
VALUES
    (@DocId, @ParentCommentId, @ThreadRootId, @Depth,
     @Body, @HasAttachment, 0,
     @CreatedBy, SYSUTCDATETIME(), SYSUTCDATETIME());";

                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                cmd.Parameters.Add(new SqlParameter("@ParentCommentId", SqlDbType.Int)
                { Value = (object?)parentId ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@ThreadRootId", SqlDbType.Int) { Value = threadRootId });
                cmd.Parameters.Add(new SqlParameter("@Depth", SqlDbType.Int) { Value = depth });
                cmd.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = body });
                cmd.Parameters.Add(new SqlParameter("@HasAttachment", SqlDbType.Bit) { Value = hasAttachment });
                cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 200) { Value = userName ?? "" });

                var newIdObj = await cmd.ExecuteScalarAsync();
                var newId = (newIdObj is int i) ? i : Convert.ToInt32(newIdObj);

                if (threadRootId == 0 && newId > 0)
                {
                    var u = conn.CreateCommand();
                    u.CommandText = @"UPDATE dbo.DocumentComments SET ThreadRootId = @id WHERE CommentId = @id;";
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = newId });
                    await u.ExecuteNonQueryAsync();
                }

                return Json(new { ok = true, id = newId });
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "PostComment failed");
                return StatusCode(500, new
                {
                    messages = new[] { "DOC_Err_RequestFailed" },
                    detail = ex.Message
                });
            }
        }

        [HttpDelete("Comments/{id}")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> DeleteComment(int id, string docId)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var userName = User?.Identity?.Name ?? "";
            var isAdmin = User?.IsInRole("Admin") ?? false;

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // 2025.12.04 Added 자식 댓글이 있는지 먼저 확인
            var checkChild = conn.CreateCommand();
            checkChild.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentComments
WHERE ParentCommentId = @id AND DocId = @docId AND IsDeleted = 0;";
            checkChild.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = id });
            checkChild.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = docId });

            var childCountObj = await checkChild.ExecuteScalarAsync();
            var childCount = (childCountObj is int ci) ? ci : Convert.ToInt32(childCountObj);
            if (childCount > 0)
            {
                // 하위 답글이 있으면 삭제 불가
                return StatusCode(409, new
                {
                    ok = false,
                    messages = new[] { "DOC_Err_DeleteFailed" },
                    detail = "comment has child replies"
                });
            }

            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
UPDATE dbo.DocumentComments
SET IsDeleted = 1, UpdatedAt = SYSUTCDATETIME()
WHERE CommentId = @id
  AND DocId = @docId
  AND IsDeleted = 0
  AND (@IsAdmin = 1 OR CreatedBy = @UserName);";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = id });
            cmd.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@IsAdmin", SqlDbType.Bit) { Value = isAdmin ? 1 : 0 });
            cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200) { Value = userName ?? "" });

            var rows = await cmd.ExecuteNonQueryAsync();
            if (rows == 0)
                return Forbid();

            return Json(new { ok = true });
        }

        [HttpPut("Comments/{id}")]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> UpdateComment(int id, [FromBody] DocCommentPostDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.body))
                return BadRequest(new { ok = false, message = "DOC_Val_Required" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var userName = User?.Identity?.Name ?? "";

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // 2025.12.04 Added 자식 댓글이 있는 경우 수정 차단
            var checkChild = conn.CreateCommand();
            checkChild.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentComments
WHERE ParentCommentId = @Id AND DocId = @DocId AND IsDeleted = 0;";
            checkChild.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int) { Value = id });
            checkChild.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId ?? "" });

            var childCountObj = await checkChild.ExecuteScalarAsync();
            var childCount = (childCountObj is int ci) ? ci : Convert.ToInt32(childCountObj);
            if (childCount > 0)
            {
                return StatusCode(409, new
                {
                    ok = false,
                    messages = new[] { "DOC_Err_RequestFailed" },
                    detail = "comment has child replies"
                });
            }

            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
UPDATE dbo.DocumentComments
SET Body = @Body,
    UpdatedAt = SYSUTCDATETIME()
WHERE CommentId = @Id
  AND DocId = @DocId
  AND IsDeleted = 0
  AND CreatedBy = @UserName;";
            cmd.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = dto.body.Trim() });
            cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = id });
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId ?? "" });
            cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200) { Value = userName });

            var rows = await cmd.ExecuteNonQueryAsync();
            if (rows == 0)
                return Forbid();

            return Json(new { ok = true });
        }
    }
}
