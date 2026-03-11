// 2026.03.10 Changed: DX Compose 컨트롤러에서 DevExpress Workbook 문서 API 사용을 제거하고 매핑 셀 강조색 날짜포맷 줄바꿈 잠금만 ClosedXML로 적용하도록 수정함
using ClosedXML.Excel;
using DevExpress.AspNetCore.Spreadsheet;
using DevExpress.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
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
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services;
using static WebApplication1.Controllers.DocControllerHelper;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class ComposeDXController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IConfiguration _cfg;
        private readonly IAntiforgery _antiforgery;
        private readonly IWebHostEnvironment _env;
        private readonly ApplicationDbContext _db;
        private readonly IDocTemplateService _tpl;
        private readonly ILogger _log;
        private readonly IEmailSender _emailSender;
        private readonly SmtpOptions _smtpOpt;
        private static readonly object _attachSeqLock = new();
        private static readonly string[] _msgUploadFailed = new[] { "DOC_Err_UploadFailed" };
        private readonly IWebPushNotifier _webPushNotifier;
        private readonly DocControllerHelper _helper;

        public ComposeDXController(
            IStringLocalizer<SharedResource> S,
            IConfiguration cfg,
            IAntiforgery antiforgery,
            IWebHostEnvironment env,
            ApplicationDbContext db,
            IDocTemplateService tpl,
            IEmailSender emailSender,
            IOptions<SmtpOptions> smtpOptions,
            ILogger<ComposeDXController> log,
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
            _helper = new DocControllerHelper(_cfg, _env, () => User);
        }

        [HttpGet("CreateDX")]
        public async Task<IActionResult> CreateDX(string templateCode)
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
                        previewJson = rebuilt;
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
                    previewJson = metaPreview;
            }

            try
            {
                var langCode = "ko";
                var orgNodes = await _helper.BuildOrgTreeNodesAsync(langCode);
                ViewBag.OrgTreeNodes = orgNodes;
            }
            catch
            {
                ViewBag.OrgTreeNodes = Array.Empty<OrgTreeNode>();
            }

            descriptorJson = BuildDescriptorJsonWithFlowGroups(templateCode, descriptorJson);

            var composeDxExcelPath = excelAbsPath;
            try
            {
                if (!string.IsNullOrWhiteSpace(excelAbsPath) && System.IO.File.Exists(excelAbsPath))
                {
                    composeDxExcelPath = await BuildComposeDxEditableWorkbookAsync(
                        templateCode,
                        excelAbsPath,
                        descriptorJson
                    );
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "CreateDX editable workbook build failed. fallback original path. templateCode={templateCode}", templateCode);
                composeDxExcelPath = excelAbsPath;
            }

            ViewBag.DescriptorJson = descriptorJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.TemplateTitle = meta.templateTitle ?? string.Empty;
            ViewBag.TemplateCode = templateCode;
            ViewBag.ExcelPath = composeDxExcelPath;
            ViewBag.DxCallbackUrl = "/Doc/dx-callback";
            ViewBag.DxDocumentId = $"compose_{Guid.NewGuid():N}";
            ViewBag.HideCellPicker = true;

            return View("~/Views/Doc/ComposeDX.cshtml");

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
                catch
                {
                    return false;
                }

                return false;
            }
        }

        [HttpGet("dx-callback")]
        [HttpPost("dx-callback")]
        public IActionResult DxCallback()
        {
            return SpreadsheetRequestProcessor.GetResponse(HttpContext);
        }

        private sealed class ComposeDxEditableField
        {
            public string Key { get; set; } = string.Empty;
            public string Type { get; set; } = "Text";
            public string A1 { get; set; } = string.Empty;
        }

        private async Task<string> BuildComposeDxEditableWorkbookAsync(
    string templateCode,
    string sourceExcelFullPath,
    string descriptorJson)
        {
            var outRoot = Path.Combine(_env.ContentRootPath, "App_Data", "DocDxCompose");
            Directory.CreateDirectory(outRoot);

            var outPath = Path.Combine(
                outRoot,
                $"{templateCode}_{DateTime.Now:yyyyMMddHHmmssfff}_{Guid.NewGuid():N}.xlsx"
            );

            System.IO.File.Copy(sourceExcelFullPath, outPath, true);

            var editableFields = ParseEditableFields(descriptorJson);
            if (editableFields.Count == 0)
                return outPath;

            using (var wb = new XLWorkbook(outPath))
            {
                var ws = wb.Worksheets.FirstOrDefault();
                if (ws != null)
                {
                    foreach (var f in editableFields)
                    {
                        if (string.IsNullOrWhiteSpace(f.A1))
                            continue;

                        try
                        {
                            var range = ws.Range(f.A1);

                            // 편집 가능 셀 표시만 유지
                            range.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF6FF");

                            if (string.Equals(f.Type, "Date", StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    var firstCell = range.FirstCell();
                                    if (firstCell != null)
                                    {
                                        firstCell.Value = DateTime.Today;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _log.LogDebug(ex, "ComposeDX date init skipped a1={a1}", f.A1);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _log.LogDebug(ex, "ComposeDX unlock/style skipped a1={a1}", f.A1);
                        }
                    }
                }

                wb.SaveAs(outPath);
            }

            await Task.CompletedTask;
            return outPath;
        }
        private static List<ComposeDxEditableField> ParseEditableFields(string descriptorJson)
        {
            var list = new List<ComposeDxEditableField>();

            if (string.IsNullOrWhiteSpace(descriptorJson))
                return list;

            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                var root = doc.RootElement;

                if (!root.TryGetProperty("inputs", out var inputs) || inputs.ValueKind != JsonValueKind.Array)
                    return list;

                foreach (var item in inputs.EnumerateArray())
                {
                    var key = item.TryGetProperty("key", out var k) && k.ValueKind == JsonValueKind.String
                        ? (k.GetString() ?? string.Empty).Trim()
                        : string.Empty;

                    var type = item.TryGetProperty("type", out var t) && t.ValueKind == JsonValueKind.String
                        ? (t.GetString() ?? "Text").Trim()
                        : "Text";

                    var a1 = item.TryGetProperty("a1", out var a) && a.ValueKind == JsonValueKind.String
                        ? (a.GetString() ?? string.Empty).Trim().ToUpperInvariant()
                        : string.Empty;

                    if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(a1))
                        continue;

                    list.Add(new ComposeDxEditableField
                    {
                        Key = key,
                        Type = type,
                        A1 = a1
                    });
                }
            }
            catch
            {
                return new List<ComposeDxEditableField>();
            }

            return list
                .GroupBy(x => $"{x.Key}||{x.A1}", StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
        }

        [HttpPost("CreateDX")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> CreateDX([FromBody] ComposePostDto dto)
        {
            ComposePostDto? resolvedDto = dto;

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
                        resolvedDto = JsonSerializer.Deserialize<ComposePostDto>(
                            json,
                            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "CreateDX: form-data dto parse failed");
                }
            }

            dto = resolvedDto!;

            if (dto is null || string.IsNullOrWhiteSpace(dto.TemplateCode))
            {
                return BadRequest(new
                {
                    messages = new[] { "DOC_Val_TemplateRequired" },
                    stage = "arg",
                    detail = "templateCode null/empty"
                });
            }

            var tc = dto.TemplateCode!.Trim();
            var inputsMap = dto.Inputs ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var userCompDept = _helper.GetUserCompDept();
            var compCd = string.IsNullOrWhiteSpace(userCompDept.compCd) ? "0000" : userCompDept.compCd!;
            var deptIdPart = string.IsNullOrWhiteSpace(userCompDept.departmentId) ? "0" : userCompDept.departmentId!;

            var (descriptorJson, _previewJsonFromTpl, title, versionId, excelPathRaw) = await _tpl.LoadMetaAsync(tc);

            excelPathRaw = NormalizeTemplateExcelPath(excelPathRaw);
            var excelPath = _helper.ToContentRootAbsolute(excelPathRaw);

            if (versionId <= 0 || string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
            {
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_TemplateNotReady" },
                    stage = "generate",
                    detail = $"Excel not found or version invalid. versionId={versionId}, path='{excelPathRaw}' -> '{excelPath}'"
                });
            }

            string tempExcelFullPath;
            try
            {
                _log.LogInformation(
                    "CreateDX GenerateExcel start tc={tc} ver={ver} path={path} inputs={cnt}",
                    tc, versionId, excelPath, inputsMap.Count);

                tempExcelFullPath = await GenerateExcelFromInputsAsync(
                    versionId,
                    excelPath,
                    inputsMap,
                    dto.DescriptorVersion);

                _log.LogInformation("CreateDX GenerateExcel done out={out}", tempExcelFullPath);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "CreateDX Generate failed tc={tc} ver={ver} path={path}", tc, versionId, excelPath);
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_SaveFailed" },
                    stage = "generate",
                    detail = ex.Message
                });
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
                toEmails = new List<string>();
                diag = new List<string> { "recipients resolve error" };
            }

            var approvalsJson = ExtractApprovalsJson(normalizedDesc);

            bool approvalsHasAnyValue = false;
            bool approvalsParsedOk = false;

            if (!string.IsNullOrWhiteSpace(approvalsJson))
            {
                try
                {
                    using var aj = JsonDocument.Parse(approvalsJson);
                    if (aj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        approvalsParsedOk = true;
                        foreach (var a in aj.RootElement.EnumerateArray())
                        {
                            string? v = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String) v = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String) v = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String) v = v3.GetString();

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
                    {
                        return BadRequest(new
                        {
                            messages = new[] { "DOC_Err_SaveFailed" },
                            stage = "approvals",
                            detail = "No valid approver emails."
                        });
                    }

                    approvalsJson = JsonSerializer.Serialize(built);
                }
                else
                {
                    return BadRequest(new
                    {
                        messages = new[] { "DOC_Err_SaveFailed" },
                        stage = "approvals",
                        detail = "Approver not resolved (toEmails empty)."
                    });
                }
            }
            else
            {
                try
                {
                    using var aj = JsonDocument.Parse(approvalsJson);
                    if (aj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        var list = new List<Dictionary<string, object>>();
                        int seq = 1;

                        foreach (var a in aj.RootElement.EnumerateArray())
                        {
                            var item = new Dictionary<string, object>();

                            string at = "Person";
                            if (a.TryGetProperty("approverType", out var at1) && at1.ValueKind == JsonValueKind.String)
                                at = string.IsNullOrWhiteSpace(at1.GetString()) ? at : at1.GetString()!;
                            else if (a.TryGetProperty("ApproverType", out var at2) && at2.ValueKind == JsonValueKind.String)
                                at = string.IsNullOrWhiteSpace(at2.GetString()) ? at : at2.GetString()!;

                            bool required = false;
                            if (a.TryGetProperty("required", out var r1) &&
                                (r1.ValueKind == JsonValueKind.True || r1.ValueKind == JsonValueKind.False))
                            {
                                required = r1.GetBoolean();
                            }

                            string? v = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String) v = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String) v = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String) v = v3.GetString();

                            item["roleKey"] = $"A{seq}";
                            item["approverType"] = at;
                            item["required"] = required;
                            item["value"] = v ?? string.Empty;

                            list.Add(item);
                            seq++;
                        }

                        approvalsJson = JsonSerializer.Serialize(list);
                    }
                }
                catch
                {
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
                    userId: User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty,
                    userName: User?.Identity?.Name ?? string.Empty,
                    compCd: compCd,
                    departmentId: deptIdPart);

                docId = pair.docId;
                outputPathForDb = pair.outputPath;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "CreateDX DB save failed with collision guard docId={docId} tc={tc}", docId, tc);

                object? sqlDiag = null;
                if (ex is SqlException se)
                {
                    sqlDiag = new { se.Number, se.State, se.Procedure, se.LineNumber };
                }

                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_SaveFailed" },
                    stage = "db",
                    detail = ex.Message,
                    sql = sqlDiag
                });
            }

            string finalExcelFullPath = string.Empty;
            try
            {
                finalExcelFullPath = _helper.ToContentRootAbsolute(outputPathForDb);

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
                        _log.LogWarning("CreateDX final excel already exists. skip move. docId={docId} final={final}", docId, finalExcelFullPath);
                    }
                }

                if (!string.IsNullOrWhiteSpace(finalExcelFullPath) && System.IO.File.Exists(finalExcelFullPath))
                    filledPreviewJson = BuildPreviewJsonFromExcel(finalExcelFullPath);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "CreateDX move generated excel failed temp={temp} final={final} docId={docId}",
                    tempExcelFullPath, finalExcelFullPath, docId);
            }

            try
            {
                await EnsureApprovalsAndSyncAsync(docId, tc);
                await FillDocumentApprovalsFromEmailsAsync(docId, toEmails);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "CreateDX EnsureApprovalsAndSync failed docId={docId} tc={tc}", docId, tc);
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
                _log.LogWarning(ex, "CreateDX Post-fix UserId update failed docId={docId}", docId);
            }

            try
            {
                var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

                var selected = (dto.SelectedRecipientUserIds ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (selected.Count > 0)
                {
                    var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
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

                    _log.LogInformation("CreateDX Shares saved docId={docId} count={cnt}", docId, selected.Count);
                }
                else
                {
                    _log.LogInformation("CreateDX No shares selected docId={docId}", docId);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "CreateDX Share save failed docId={docId}", docId);
            }

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;

                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                var a1UserId = await DocControllerHelper.GetA1ApproverUserIdByDocAsync(conn, docId);
                if (!string.IsNullOrWhiteSpace(a1UserId))
                {
                    await DocControllerHelper.SendApprovalPendingBadgeAsync(
                        notifier: _webPushNotifier,
                        S: _S,
                        conn: conn,
                        targetUserIds: new List<string> { a1UserId.Trim() },
                        url: "/",
                        tag: "badge-approval-pending");
                }
                else
                {
                    _log.LogWarning("CreateDX WebPush A1 userId not found docId={docId}", docId);
                }

                var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
                var shareIds = (dto.SelectedRecipientUserIds ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (shareIds.Count > 0)
                {
                    await DocControllerHelper.SendSharedUnreadBadgeAsync(
                        notifier: _webPushNotifier,
                        S: _S,
                        conn: conn,
                        targetUserIds: shareIds,
                        url: "/",
                        tag: "badge-shared");
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "CreateDX WebPush notify failed docId={docId}", docId);
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
                uploadUrl
            });
        }

        private async Task<string> GenerateExcelFromInputsAsync(long templateVersionId, string templateExcelFullPath, Dictionary<string, string> inputs, string? descriptorVersion)
        {
            if (string.IsNullOrWhiteSpace(templateExcelFullPath) || !System.IO.File.Exists(templateExcelFullPath))
                throw new FileNotFoundException("Template excel not found", templateExcelFullPath);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            var maps = new List<(string key, string a1, string type)>();

            await using (var conn = new SqlConnection(cs))
            {
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT [Key], [Type], A1, CellRow, CellColumn
FROM dbo.DocTemplateField
WHERE VersionId = @vid
ORDER BY Id;";
                cmd.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = templateVersionId });

                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    var key = rd["Key"] as string ?? string.Empty;
                    var typ = rd["Type"] as string ?? "Text";
                    string a1 = rd["A1"] as string ?? string.Empty;

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

            using var wb = new XLWorkbook(templateExcelFullPath);
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

                if (raw.Contains('\n') || raw.Contains('\r'))
                    cell.Style.Alignment.WrapText = true;

                switch (m.type.ToLowerInvariant())
                {
                    case "date":
                        if (DateTime.TryParse(raw, out var dt))
                        {
                            cell.Value = dt;
                            cell.Style.DateFormat.Format = GetExcelDateFormatByCulture(CultureInfo.CurrentUICulture);
                        }
                        else
                        {
                            cell.Value = raw;
                        }
                        break;

                    case "num":
                        if (decimal.TryParse(raw, out var dec))
                            cell.Value = dec;
                        else
                            cell.Value = raw;
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
                if (row < 1 || col < 1) return string.Empty;

                string letters = string.Empty;
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
                var s = (t ?? string.Empty).Trim().ToLowerInvariant();
                if (s.StartsWith("date")) return "date";
                if (s.StartsWith("num") || s.Contains("number") || s.Contains("decimal") || s.Contains("integer")) return "num";
                return "text";
            }
        }

        private static string GetExcelDateFormatByCulture(CultureInfo culture)
        {
            var name = (culture?.Name ?? string.Empty).Trim().ToLowerInvariant();

            return name switch
            {
                "ko-kr" => "yyyy-mm-dd",
                "vi-vn" => "dd/mm/yyyy",
                "en-us" => "mm/dd/yyyy",
                "id-id" => "dd/mm/yyyy",
                "zh-cn" => "yyyy-mm-dd",
                _ => "yyyy-mm-dd"
            };
        }
        private string BuildDescriptorJsonWithFlowGroups(string templateCode, string rawDescriptorJson)
        {
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

            desc.Approvals = RebuildApprovalsFromLegacyDescriptor(rawDescriptorJson, desc.Approvals);
            desc.FlowGroups = groups;

            var json = JsonSerializer.Serialize(
                desc,
                new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                    WriteIndented = false
                });

            return DedupApprovalsJsonByStep(json);
        }

        private static string ConvertDescriptorIfNeeded(string? json)
        {
            const string EMPTY = "{\"inputs\":[],\"approvals\":[],\"version\":\"converted\"}";
            if (!TryParseJsonFlexible(json, out var doc)) return EMPTY;

            try
            {
                var root = doc.RootElement;
                if (root.TryGetProperty("inputs", out _)) return json!;

                var inputs = new List<(string key, string type, string a1)>();
                if (root.TryGetProperty("Fields", out var fieldsEl) && fieldsEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var f in fieldsEl.EnumerateArray())
                    {
                        var key = f.TryGetProperty("Key", out var k) ? (k.GetString() ?? string.Empty).Trim() : string.Empty;
                        if (string.IsNullOrEmpty(key)) continue;

                        var type = f.TryGetProperty("Type", out var t) ? (t.GetString() ?? "Text") : "Text";
                        string a1 = string.Empty;

                        if (f.TryGetProperty("Cell", out var cell) && cell.ValueKind == JsonValueKind.Object)
                            a1 = cell.TryGetProperty("A1", out var a) ? (a.GetString() ?? string.Empty) : string.Empty;

                        inputs.Add((key, MapType(type), a1));
                    }
                }

                var approvals = new List<object>();
                if (root.TryGetProperty("Approvals", out var apprEl) && apprEl.ValueKind == JsonValueKind.Array)
                {
                    int index = 0;
                    foreach (var a in apprEl.EnumerateArray())
                    {
                        var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                        var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;

                        approvals.Add(new
                        {
                            roleKey = $"A{index + 1}",
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
                t = (t ?? string.Empty).Trim().ToLowerInvariant();
                if (t.StartsWith("date")) return "Date";
                if (t.StartsWith("num") || t.Contains("number") || t.Contains("decimal") || t.Contains("integer")) return "Num";
                return "Text";
            }

            static string MapApproverType(string t)
            {
                t = (t ?? string.Empty).Trim();
                return (t == "Person" || t == "Role" || t == "Rule") ? t : "Person";
            }
        }

        private static List<object>? RebuildApprovalsFromLegacyDescriptor(string? legacyJson, List<object>? current)
        {
            if (!TryParseJsonFlexible(legacyJson, out var doc))
                return current;

            try
            {
                var root = doc.RootElement;
                if (!root.TryGetProperty("Approvals", out var apprEl) || apprEl.ValueKind != JsonValueKind.Array)
                    return current;

                var list = new List<object>();
                int index = 1;

                foreach (var a in apprEl.EnumerateArray())
                {
                    var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                    var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;
                    var mappedType = (typ == "Person" || typ == "Role" || typ == "Rule") ? typ : "Person";

                    list.Add(new
                    {
                        roleKey = $"A{index}",
                        approverType = mappedType,
                        required = false,
                        value = val ?? string.Empty
                    });

                    index++;
                }

                if (list.Count == 0) return current;
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

        private static string DedupApprovalsJsonByStep(string json)
        {
            try
            {
                var root = System.Text.Json.Nodes.JsonNode.Parse(json) as System.Text.Json.Nodes.JsonObject;
                if (root == null) return json;

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
                        var m = Regex.Match(rk.Trim(), @"(?i)^A(\d+)$");
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

        private long GetLatestVersionId(string templateCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
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
            catch
            {
            }

            return 0;
        }

        private List<FlowGroupDto> BuildFlowGroupsForTemplate(long versionId)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
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
                    var key = rd["Key"] as string ?? string.Empty;
                    var a1 = rd["A1"] as string;
                    var r = rd["CellRow"] is DBNull ? 0 : Convert.ToInt32(rd["CellRow"]);
                    var c = rd["CellColumn"] is DBNull ? 0 : Convert.ToInt32(rd["CellColumn"]);
                    var (baseKey, idx) = ParseKey(key);
                    rows.Add((key, a1, r, c, baseKey, idx));
                }
            }

            return rows
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
        }

        private static string ExtractApprovalsJson(string descriptorJson)
        {
            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                if (doc.RootElement.TryGetProperty("approvals", out var arr)) return arr.GetRawText();
                if (doc.RootElement.TryGetProperty("Approvals", out var arr2)) return arr2.GetRawText();
            }
            catch
            {
            }

            return "[]";
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
                    if (LooksLikeEmail(tk))
                    {
                        emails.Add(tk);
                        diag.Add($"'{tk}' -> email");
                        continue;
                    }

                    if (conn == null)
                    {
                        diag.Add($"'{tk}' -> no-conn");
                        continue;
                    }

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
                    {
                        emails.Add(email!);
                        diag.Add($"'{tk}' -> AspNetUsers '{email}'");
                        continue;
                    }

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
                    {
                        emails.Add(email!);
                        diag.Add($"'{tk}' -> UserProfiles '{email}'");
                    }
                    else
                    {
                        diag.Add($"'{tk}' not resolved");
                    }
                }
            }

            var cs = cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            using var conn = string.IsNullOrWhiteSpace(cs) ? null : new SqlConnection(cs);
            if (conn != null) conn.Open();

            if (dto?.Mail?.TO is { Count: > 0 })
                foreach (var s in dto.Mail.TO) AppendTokens(s, conn);

            if (emails.Count == 0 && dto?.Approvals?.To is { Count: > 0 })
                foreach (var s in dto.Approvals.To) AppendTokens(s, conn);

            if (emails.Count == 0 && dto?.Approvals?.Steps is { Count: > 0 })
            {
                foreach (var st in dto.Approvals.Steps)
                {
                    if (!string.IsNullOrWhiteSpace(st?.Value))
                        AppendTokens(st!.Value!, conn);
                }
            }

            if (emails.Count == 0)
            {
                try
                {
                    using var doc = JsonDocument.Parse(normalizedDesc);
                    if (doc.RootElement.TryGetProperty("approvals", out var arr) && arr.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var el in arr.EnumerateArray())
                        {
                            var val = el.TryGetProperty("value", out var v) ? v.GetString()
                                     : el.TryGetProperty("ApproverValue", out var v2) ? v2.GetString()
                                     : null;

                            if (!string.IsNullOrWhiteSpace(val))
                                AppendTokens(val!, conn);
                        }
                    }
                }
                catch
                {
                }
            }

            if (emails.Count == 0 && !string.IsNullOrWhiteSpace(docId) && conn != null)
            {
                try
                {
                    using var q = conn.CreateCommand();
                    q.CommandText = @"
SELECT ApproverValue
FROM dbo.DocumentApprovals
WHERE DocId=@id
  AND ApproverValue IS NOT NULL
  AND LTRIM(RTRIM(ApproverValue))<>'';";
                    q.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });

                    using var rd = q.ExecuteReader();
                    while (rd.Read()) AppendTokens(rd.GetString(0), conn);
                    diag.Add("fallback: DocumentApprovals used");
                }
                catch (Exception ex)
                {
                    diag.Add("db fallback ex: " + ex.Message);
                }
            }

            if (emails.Count == 0 && !string.IsNullOrWhiteSpace(templateCode) && conn != null)
            {
                try
                {
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = @"
SELECT TOP (1) v.DescriptorJson
FROM dbo.DocTemplateVersion v
JOIN dbo.DocTemplateMaster m ON m.Id=v.TemplateId
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
                                if (!string.IsNullOrWhiteSpace(val))
                                    AppendTokens(val!, conn);
                            }
                            diag.Add("fallback: Template.DescriptorJson approvals used");
                        }
                    }
                }
                catch (Exception ex)
                {
                    diag.Add("template fallback ex: " + ex.Message);
                }
            }

            return (emails.Distinct(StringComparer.OrdinalIgnoreCase).ToList(), diag);
        }

        private async Task FillDocumentApprovalsFromEmailsAsync(string docId, IEnumerable<string>? approverEmails)
        {
            var list = approverEmails?
                .Where(e => !string.IsNullOrWhiteSpace(e))
                .Select(e => e.Trim())
                .ToList() ?? new List<string>();

            _log.LogInformation("CreateDX FillDocumentApprovalsFromEmailsAsync START docId={docId}, rawEmails=[{emails}]",
                docId, string.Join(", ", list));

            if (list.Count == 0)
            {
                _log.LogWarning("CreateDX FillDocumentApprovalsFromEmailsAsync docId={docId} email list empty.", docId);
                return;
            }

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            for (var i = 0; i < list.Count; i++)
            {
                var step = i + 1;
                var email = list[i];

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
                    cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)userId ?? DBNull.Value });

                    await cmd.ExecuteNonQueryAsync();
                }
            }

            _log.LogInformation("CreateDX FillDocumentApprovalsFromEmailsAsync END docId={docId}", docId);
        }

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
                        docId,
                        templateCode,
                        title,
                        status,
                        outputPath,
                        inputs,
                        approvalsJson,
                        descriptorJson,
                        userId,
                        userName,
                        compCd,
                        departmentId);

                    return (docId, outputPath);
                }
                catch (SqlException se) when (se.Number == 2627 || se.Number == 2601)
                {
                    _log.LogWarning(se, "CreateDX DocId unique violation attempt={attempt} docId={docId}", attempt + 1, docId);

                    var newDocId = WithNewSuffix(docId);
                    var newPath = EnsurePathFor(outputPath, newDocId);

                    if (System.IO.File.Exists(newPath))
                    {
                        newDocId = WithNewSuffix(newDocId);
                        newPath = EnsurePathFor(outputPath, newDocId);
                    }

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(outputPath)
                            && System.IO.File.Exists(outputPath)
                            && !string.Equals(outputPath, newPath, StringComparison.OrdinalIgnoreCase))
                        {
                            System.IO.File.Move(outputPath, newPath, overwrite: false);
                        }
                    }
                    catch (Exception exMove)
                    {
                        _log.LogWarning(exMove, "CreateDX Document file rename failed old={old} new={new}", outputPath, newPath);
                    }

                    docId = newDocId;
                    outputPath = newPath;
                }
            }

            throw new InvalidOperationException("DocId collision could not be resolved after retries");
        }

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
            int? templateVersionId = null)
        {
            if (string.IsNullOrWhiteSpace(compCd))
                throw new InvalidOperationException("CompCd is required by schema but missing.");

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            await using var tx = await conn.BeginTransactionAsync();

            try
            {
                int? effectiveTemplateVersionId = templateVersionId;

                if (!effectiveTemplateVersionId.HasValue)
                {
                    await using var vCmd = conn.CreateCommand();
                    vCmd.Transaction = (SqlTransaction)tx;
                    vCmd.CommandText = @"
SELECT TOP (1) v.Id
FROM dbo.DocTemplateVersion AS v
JOIN dbo.DocTemplateMaster AS m ON v.TemplateId = m.Id
WHERE m.DocCode = @TemplateCode
ORDER BY v.VersionNo DESC, v.Id DESC;";
                    vCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode });

                    var obj = await vCmd.ExecuteScalarAsync();
                    if (obj != null && obj != DBNull.Value && int.TryParse(obj.ToString(), out var vid))
                        effectiveTemplateVersionId = vid;
                }

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
                    cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 450) { Value = userId ?? string.Empty });
                    cmd.Parameters.Add(new SqlParameter("@CreatedByName", SqlDbType.NVarChar, 200) { Value = userName ?? string.Empty });
                    cmd.Parameters.Add(new SqlParameter("@DescriptorJson", SqlDbType.NVarChar, -1) { Value = (object?)descriptorJson ?? DBNull.Value });
                    await cmd.ExecuteNonQueryAsync();
                }

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
                            var m = Regex.Match(roleKey.Trim(), @"(?i)^A(\d+)$");
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

                        string roleKey = string.Empty;
                        if (el.TryGetProperty("roleKey", out var rk) && rk.ValueKind == JsonValueKind.String) roleKey = rk.GetString() ?? string.Empty;
                        else if (el.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;

                        var step = GetStep(el, roleKey);
                        if (step <= 0) step = 1;

                        var rkNorm = string.IsNullOrWhiteSpace(roleKey) ? $"A{step}" : roleKey.Trim();
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
                }

                await tx.CommitAsync();
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }
        }

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

        private async Task EnsureApprovalsAndSyncAsync(string docId, string docCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();
            await using var tx = await conn.BeginTransactionAsync();

            try
            {
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

                if (string.IsNullOrWhiteSpace(descriptorJson))
                {
                    await tx.CommitAsync();
                    return;
                }

                var approvals = new List<(int StepOrder, string RoleKey, string? ApproverValue)>();

                try
                {
                    using var dj = JsonDocument.Parse(descriptorJson);
                    var root = dj.RootElement;

                    if (root.TryGetProperty("approvals", out var arr) && arr.ValueKind == JsonValueKind.Array)
                    {
                        var i = 0;
                        foreach (var a in arr.EnumerateArray())
                        {
                            string? approverValue = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String) approverValue = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String) approverValue = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String) approverValue = v3.GetString();

                            if (string.IsNullOrWhiteSpace(approverValue))
                            {
                                i++;
                                continue;
                            }

                            int stepOrder = i + 1;
                            if (a.TryGetProperty("order", out var ordEl))
                            {
                                if (ordEl.ValueKind == JsonValueKind.Number && ordEl.TryGetInt32(out var ordNum)) stepOrder = ordNum;
                                else if (ordEl.ValueKind == JsonValueKind.String && int.TryParse(ordEl.GetString(), out var ordStr)) stepOrder = ordStr;
                            }
                            else if (a.TryGetProperty("slot", out var slotEl))
                            {
                                if (slotEl.ValueKind == JsonValueKind.Number && slotEl.TryGetInt32(out var slotNum)) stepOrder = slotNum;
                                else if (slotEl.ValueKind == JsonValueKind.String && int.TryParse(slotEl.GetString(), out var slotStr)) stepOrder = slotStr;
                            }

                            string roleKey = string.Empty;
                            if (a.TryGetProperty("roleKey", out var rk1) && rk1.ValueKind == JsonValueKind.String) roleKey = rk1.GetString() ?? string.Empty;
                            else if (a.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;
                            else if (a.TryGetProperty("part", out var rk3) && rk3.ValueKind == JsonValueKind.String) roleKey = rk3.GetString() ?? string.Empty;

                            if (string.IsNullOrWhiteSpace(roleKey))
                                roleKey = $"A{stepOrder}";

                            approvals.Add((stepOrder, roleKey, approverValue!.Trim()));
                            i++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "CreateDX EnsureApprovalsAndSyncAsync parse failed docId={docId}", docId);
                }

                if (approvals.Count == 0)
                {
                    await tx.CommitAsync();
                    return;
                }

                await using (var del = conn.CreateCommand())
                {
                    del.Transaction = (SqlTransaction)tx;
                    del.CommandText = @"DELETE FROM dbo.DocumentApprovals WHERE DocId = @DocId;";
                    del.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    await del.ExecuteNonQueryAsync();
                }

                var ordered = approvals
                    .OrderBy(a => a.StepOrder)
                    .ThenBy(a => a.RoleKey, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                for (int idx = 0; idx < ordered.Count; idx++)
                {
                    var a = ordered[idx];
                    var stepOrder = idx + 1;
                    var roleKey = $"A{stepOrder}";

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
                    ins.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128) { Value = a.ApproverValue! });

                    await ins.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();

                _log.LogInformation(
                    "CreateDX EnsureApprovalsAndSyncAsync completed docId={docId}, docCode={docCode}, count={cnt}",
                    docId, docCode, ordered.Count);
            }
            catch (Exception ex)
            {
                try { await tx.RollbackAsync(); } catch { }
                _log.LogError(ex, "CreateDX EnsureApprovalsAndSyncAsync failed docId={docId}, docCode={docCode}", docId, docCode);
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

        private static bool HasCells(string json)
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
            catch
            {
            }

            return false;
        }

        private static string SafeDxDocumentId(string raw)
        {
            raw ??= string.Empty;
            raw = raw.Trim();
            if (raw.Length == 0) raw = "doc";
            raw = Regex.Replace(raw, @"[^a-zA-Z0-9_\-\.]", "_");
            if (raw.Length > 120) raw = raw[..120];
            return raw;
        }

        private static bool TryParseJsonFlexible(string? json, out JsonDocument doc)
        {
            doc = null!;
            if (string.IsNullOrWhiteSpace(json)) return false;

            try
            {
                doc = JsonDocument.Parse(json);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static (string BaseKey, int? Index) ParseKey(string key)
        {
            var s = (key ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(s))
                return (string.Empty, null);

            var m = Regex.Match(s, @"^(.*?)(?:_(\d+))?$", RegexOptions.CultureInvariant);
            if (!m.Success)
                return (s, null);

            var baseKey = (m.Groups[1].Value ?? string.Empty).Trim();
            if (int.TryParse(m.Groups[2].Value, out var n))
                return (baseKey, n);

            return (baseKey, null);
        }

        private string NormalizeTemplateExcelPath(string? raw)
        {
            var s = (raw ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(s)) return s;

            if (s.StartsWith("App_Data\\", StringComparison.OrdinalIgnoreCase) ||
                s.StartsWith("App_Data/", StringComparison.OrdinalIgnoreCase))
                return s.Replace('/', '\\');

            var idx = s.IndexOf("App_Data\\", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) idx = s.IndexOf("App_Data/", StringComparison.OrdinalIgnoreCase);

            return idx >= 0 ? s[idx..].Replace('/', '\\') : s;
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

            [JsonPropertyName("selectedRecipientUserIds")]
            public List<string>? SelectedRecipientUserIds { get; set; }

            [JsonPropertyName("attachments")]
            public List<ComposeAttachmentDto>? Attachments { get; set; }
        }

        public sealed class ComposeAttachmentDto
        {
            [JsonPropertyName("FileKey")]
            public string? FileKey { get; set; }

            [JsonPropertyName("OriginalName")]
            public string? OriginalName { get; set; }

            [JsonPropertyName("ContentType")]
            public string? ContentType { get; set; }

            [JsonPropertyName("ByteSize")]
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
            public string Key { get; set; } = string.Empty;
            public string? A1 { get; set; }
            public string? Type { get; set; }
        }

        private sealed class FlowGroupDto
        {
            public string ID { get; set; } = string.Empty;
            public List<string> Keys { get; set; } = new();
        }
    }
}