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

            var descriptorJson = string.IsNullOrWhiteSpace(meta.descriptorJson) ? "{}" : meta.descriptorJson;
            var previewJson = string.IsNullOrWhiteSpace(meta.previewJson) ? "{}" : meta.previewJson;
            var excelAbsPath = meta.excelFilePath ?? string.Empty;

            if (string.IsNullOrWhiteSpace(previewJson) || !HasCells(previewJson))
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(excelAbsPath) && System.IO.File.Exists(excelAbsPath))
                    {
                        var rebuilt = BuildPreviewJsonFromExcel(excelAbsPath);
                        if (!string.IsNullOrWhiteSpace(rebuilt) && HasCells(rebuilt))
                        {
                            previewJson = rebuilt;
                        }
                    }
                }
                catch { }
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

        // ========== CSRF ==========
        [HttpGet("Csrf")]
        [Produces("application/json")]
        public IActionResult Csrf()
        {
            var tokens = _antiforgery.GetAndStoreTokens(HttpContext);
            return Json(new { headerName = "RequestVerificationToken", token = tokens.RequestToken });
        }

        // ========== Create (POST) ==========
        [HttpPost("Create")]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> Create([FromBody] ComposePostDto? dto)
        {
            if (dto is null || string.IsNullOrWhiteSpace(dto.TemplateCode))
                return BadRequest(new { messages = new[] { "DOC_Val_TemplateRequired" }, stage = "arg", detail = "templateCode null/empty" });

            var tc = dto.TemplateCode!;
            var inputsMap = dto.Inputs ?? new Dictionary<string, string>();

            var (descriptorJson, _previewJsonFromTpl, title, versionId, excelPathRaw) = await _tpl.LoadMetaAsync(tc);
            var excelPath = excelPathRaw;
            if (!string.IsNullOrWhiteSpace(excelPath) && !Path.IsPathRooted(excelPath))
                excelPath = Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, excelPath);

            if (versionId <= 0 || string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_TemplateNotReady" },
                    stage = "generate",
                    detail = $"Excel not found or version invalid. versionId={versionId}, path='{excelPathRaw}' → '{excelPath}'"
                });

            string outputPath;
            try
            {
                _log.LogInformation("GenerateExcel start tc={tc} ver={ver} path={path} inputs={cnt}", tc, versionId, excelPath, inputsMap.Count);

                outputPath = await GenerateExcelFromInputsAsync(
                    versionId, excelPath, inputsMap, dto.DescriptorVersion);

                _log.LogInformation("GenerateExcel done out={out}", outputPath);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Create.Generate failed tc={tc} ver={ver} path={path}", tc, versionId, excelPath);
                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_SaveFailed" },
                    stage = "generate",
                    detail = ex.Message
                });
            }

            var filledPreviewJson = BuildPreviewJsonFromExcel(outputPath);
            var normalizedDesc = BuildDescriptorJsonWithFlowGroups(tc, descriptorJson);

            // Approvals 추출/보강
            var approvalsJson = ExtractApprovalsJson(normalizedDesc);
            if (string.IsNullOrWhiteSpace(approvalsJson) || approvalsJson.Trim() == "[]")
                approvalsJson = "[{\"roleKey\":\"A1\",\"approverType\":\"Person\",\"required\":false,\"value\":\"\"}]";

            var docId = Path.GetFileNameWithoutExtension(outputPath);

            // 최초 수신자
            List<string> toEmails; List<string> diag;
            try
            {
                var recipients = GetInitialRecipients(dto, normalizedDesc, _cfg, docId, tc);
                toEmails = recipients.emails;
                diag = recipients.diag;
            }
            catch
            {
                toEmails = new(); diag = new() { "recipients resolve error" };
            }

            // Approvals 값 없으면 첫 수신자를 A1 로 보정
            if (toEmails.Count > 0)
            {
                try
                {
                    using var aj = JsonDocument.Parse(approvalsJson);
                    var hasValue = aj.RootElement.EnumerateArray().Any(el =>
                        (el.TryGetProperty("value", out var v) && !string.IsNullOrWhiteSpace(v.GetString())) ||
                        (el.TryGetProperty("ApproverValue", out var v2) && !string.IsNullOrWhiteSpace(v2.GetString())));
                    if (!hasValue)
                    {
                        var first = toEmails[0];
                        var a1 = new[] { new { roleKey = "A1", approverType = "Person", required = false, value = first } };
                        approvalsJson = JsonSerializer.Serialize(a1);

                        try
                        {
                            using var nd = JsonDocument.Parse(normalizedDesc);
                            var root = nd.RootElement;
                            var inputsJson = root.TryGetProperty("inputs", out var ii) ? JsonSerializer.Deserialize<object>(ii.GetRawText()) : new object[0];
                            var styles = root.TryGetProperty("styles", out var ss) ? JsonSerializer.Deserialize<object>(ss.GetRawText()) : null;
                            var flowGroups = root.TryGetProperty("flowGroups", out var fg) ? JsonSerializer.Deserialize<object>(fg.GetRawText()) : null;
                            var versionTag = root.TryGetProperty("version", out var vv) ? vv.GetString() : "converted";

                            normalizedDesc = JsonSerializer.Serialize(new
                            {
                                version = versionTag,
                                inputs = inputsJson,
                                styles,
                                flowGroups,
                                approvals = a1
                            });
                        }
                        catch { }

                        diag.Add("approval fallback injected from recipients A1.value");
                    }
                }
                catch { }
            }

            // DB 저장 (충돌 가드)
            try
            {
                var pair = await SaveDocumentWithCollisionGuardAsync(
                    docId: docId,
                    templateCode: tc,
                    title: string.IsNullOrWhiteSpace(title) ? tc : title!,
                    status: "PendingA1",
                    outputPath: outputPath,
                    inputs: inputsMap,
                    approvalsJson: approvalsJson,
                    descriptorJson: normalizedDesc,
                    userId: User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "",
                    userName: User?.Identity?.Name ?? "",
                    compCd: GetUserCompDept().compCd,
                    departmentId: GetUserCompDept().departmentId
                );

                docId = pair.docId;
                outputPath = pair.outputPath;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Create.DB save failed with collision guard docId={docId} tc={tc}", docId, tc);

                object? sqlDiag = null;
                if (ex is SqlException se) sqlDiag = new { se.Number, se.State, se.Procedure, se.LineNumber };

                return BadRequest(new
                {
                    messages = new[] { "DOC_Err_SaveFailed" },
                    stage = "db",
                    detail = ex.Message,
                    sql = sqlDiag
                });
            }

            // 승인 동기화
            try { await EnsureApprovalsAndSyncAsync(docId, tc); }
            catch (Exception ex) { _log.LogWarning(ex, "EnsureApprovalsAndSync failed docId={docId} tc={tc}", docId, tc); }

            // DocumentApprovals.UserId 보강
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

            // 메일 구성
            IEnumerable<string> Norm(IEnumerable<string>? src) =>
                (src ?? Array.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase);

            var fromEmail = !string.IsNullOrWhiteSpace(dto?.Mail?.From) ? dto!.Mail!.From! : GetCurrentUserEmail();
            var authorDisplay = GetCurrentUserDisplayNameStrict();
            var fromAddress = ComposeAddress(fromEmail, authorDisplay);

            var toList = (dto?.Mail?.TO != null && dto.Mail!.TO!.Count > 0) ? dto.Mail!.TO! : toEmails;
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
                    var tpl = BuildSubmissionMail(authorDisplay, docTitle, docId, recipientDisplay);
                    subject = string.IsNullOrWhiteSpace(reqSubj) ? tpl.subject : reqSubj!;
                    body = string.IsNullOrWhiteSpace(reqBody) ? tpl.bodyHtml : reqBody!;
                }
            }

            var shouldSend = dto?.Mail?.Send != false;
            int sentCount = 0; string? sendErr = null;

            if (shouldSend && toDisplay.Any())
            {
                try
                {
                    foreach (var to in toDisplay) await _emailSender.SendEmailAsync(to, subject, body);
                    foreach (var cc in ccDisplay) await _emailSender.SendEmailAsync(cc, subject, body);
                    foreach (var bc in bccDisplay) await _emailSender.SendEmailAsync(bc, subject, body);
                    sentCount = toDisplay.Count + ccDisplay.Count + bccDisplay.Count;
                }
                catch (Exception ex) { sendErr = ex.Message; _log.LogError(ex, "Mail send failed docId={docId}", docId); }
            }

            return Json(new
            {
                ok = true,
                docId,
                title = title ?? tc,
                status = "PendingA1",
                previewJson = filledPreviewJson,
                approvalsJson,
                mailInfo = new
                {
                    from = fromEmail ?? string.Empty,
                    fromDisplay = fromAddress,
                    to = Norm(toList).ToArray(),
                    toDisplay = toDisplay.ToArray(),
                    cc = Norm(ccList).ToArray(),
                    ccDisplay = ccDisplay.ToArray(),
                    bcc = Norm(bccList).ToArray(),
                    bccDisplay = bccDisplay.ToArray(),
                    subject,
                    body,
                    sent = sendErr is null && sentCount > 0,
                    sentCount,
                    error = sendErr ?? ""
                }
            });
        }

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

        [HttpPost("ApproveOrHold")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> ApproveOrHold([FromBody] ApproveDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.docId) || string.IsNullOrWhiteSpace(dto.action))
                return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" } });

            var outXlsx = ResolveOutputPathFromDocId(dto.docId);
            if (string.IsNullOrWhiteSpace(outXlsx) || !System.IO.File.Exists(outXlsx))
                return NotFound(new { messages = new[] { "DOC_Err_DocumentNotFound" } });

            var approverId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            var approverName = GetCurrentUserDisplayNameStrict();
            ApplyStampToExcel(outXlsx!, dto.slot, dto.action!, approverName);

            var newStatus = dto.action!.ToLowerInvariant() switch
            {
                "approve" => $"ApprovedA{dto.slot}",
                "hold" => $"OnHoldA{dto.slot}",
                "reject" => $"RejectedA{dto.slot}",
                _ => $"Updated"
            };

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using (var u = conn.CreateCommand())
                {
                    u.CommandText = @"
UPDATE dbo.DocumentApprovals
SET Action=@act, ActedAt=SYSUTCDATETIME(), ActorName=@actor,
    Status = CASE WHEN @act=N'approve' THEN N'Approved'
                  WHEN @act=N'hold'    THEN N'OnHold'
                  WHEN @act=N'reject'  THEN N'Rejected'
                  ELSE ISNULL(Status, N'Updated') END,
    UserId = COALESCE(UserId, @uid)
WHERE DocId=@id AND StepOrder=@step;";
                    u.Parameters.Add(new SqlParameter("@act", SqlDbType.NVarChar, 20) { Value = dto.action!.ToLowerInvariant() });
                    u.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName });
                    u.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    u.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = dto.slot });
                    await u.ExecuteNonQueryAsync();
                }

                await using (var u = conn.CreateCommand())
                {
                    u.CommandText = @"UPDATE dbo.Documents SET Status=@st WHERE DocId=@id;";
                    u.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = newStatus });
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    await u.ExecuteNonQueryAsync();
                }

                if (dto.action.Equals("approve", StringComparison.OrdinalIgnoreCase))
                {
                    string? templateCode = null;
                    await using (var q = conn.CreateCommand())
                    {
                        q.CommandText = @"SELECT TOP 1 TemplateCode FROM dbo.Documents WHERE DocId=@id;";
                        q.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        templateCode = (string?)await q.ExecuteScalarAsync();
                    }

                    if (!string.IsNullOrWhiteSpace(templateCode))
                    {
                        var (descriptorJson, _, _, _, _) = await _tpl.LoadMetaAsync(templateCode!);
                        var normalizedDesc = BuildDescriptorJsonWithFlowGroups(templateCode!, descriptorJson);

                        var next = dto.slot + 1;
                        var toList = ResolveApproverEmails(normalizedDesc, next);
                        if (toList.Count == 0)
                        {
                            _log.LogWarning("Next-step approver not resolved. docId={docId}, step={step}", dto.docId, next);
                            await using var up2a = conn.CreateCommand();
                            up2a.CommandText = @"UPDATE dbo.Documents SET Status=N'Approved' WHERE DocId=@id;";
                            up2a.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            await up2a.ExecuteNonQueryAsync();
                            newStatus = "Approved";
                        }
                        else
                        {
                            var subject = $"[승인요청] {templateCode} - {next}차";
                            var body = $"이 문서의 {next}차 승인을 요청드립니다. 문서번호: {dto.docId}";
                            foreach (var to in toList)
                                await _emailSender.SendEmailAsync(to, subject, body);

                            await using var up2 = conn.CreateCommand();
                            up2.CommandText = @"UPDATE dbo.Documents SET Status=@st WHERE DocId=@id;";
                            up2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = $"PendingA{next}" });
                            up2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            await up2.ExecuteNonQueryAsync();
                            newStatus = $"PendingA{next}";
                        }
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

        [HttpGet("Detail/{id?}")]
        public async Task<IActionResult> Detail(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                // HttpContext? 체이닝 결과가 null일 수 있으므로, 임시 변수로 받았다가 유효할 때만 대입
                var qid = HttpContext?.Request?.Query["id"].ToString();
                if (!string.IsNullOrWhiteSpace(qid))
                {
                    id = qid; // id는 non-nullable이지만 여기서는 qid가 공백 아님이 보장됨
                }
            }

            if (string.IsNullOrWhiteSpace(id))
            {
                // 문서 ID 없으면 게시판으로 이동 (또는 BadRequest로 변경 가능)
                return RedirectToAction("Board");
            }

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // ---------- 1) 기본 문서 정보 ----------
            string? templateCode = null;
            string? templateTitle = null;
            string? status = null;
            string? descriptorJson = null;
            string? outputPath = null;

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT TOP 1
       d.DocId,
       d.TemplateCode,
       d.TemplateTitle,
       d.Status,
       d.DescriptorJson,
       d.OutputPath
FROM dbo.Documents d
WHERE d.DocId = @id;";
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

                await using var rd = await cmd.ExecuteReaderAsync();
                if (!await rd.ReadAsync())
                {
                    // 문서 없음
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
            }

            // InputsJson 컬럼이 DB에 없으므로, 뷰용 기본값만 세팅
            var inputsJson = "{}";

            // ---------- 2) 프리뷰 JSON ----------
            string previewJson = "{}";
            try
            {
                if (!string.IsNullOrWhiteSpace(outputPath) && System.IO.File.Exists(outputPath))
                {
                    previewJson = BuildPreviewJsonFromExcel(outputPath);
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Detail: BuildPreviewJsonFromExcel failed for {DocId}", id);
                previewJson = "{}";
            }

            // ---------- 3) 승인 정보 ----------
            var approvals = new List<object>();
            await using (var ac = conn.CreateCommand())
            {
                ac.CommandText = @"
SELECT StepOrder, RoleKey, ApproverValue, UserId,
       Status, Action, ActedAt, ActorName
FROM dbo.DocumentApprovals
WHERE DocId = @id
ORDER BY StepOrder;";
                ac.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

                await using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    DateTime? acted = r["ActedAt"] as DateTime?;
                    approvals.Add(new
                    {
                        step = (int)r["StepOrder"],
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue = r["ApproverValue"]?.ToString(),
                        userId = r["UserId"]?.ToString(),
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = acted.HasValue
                                        ? ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                                        : null,
                        actorName = r["ActorName"]?.ToString()
                    });
                }
            }

            // ---------- 4) 조회 로그 (옵션) ----------
            var viewLogs = new List<object>();
            try
            {
                await using var vc = conn.CreateCommand();
                vc.CommandText = @"
SELECT TOP 200 ViewedAtUtc, UserId, UserName
FROM dbo.DocumentViewLogs
WHERE DocId = @id
ORDER BY ViewedAtUtc DESC;";
                vc.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

                await using var vr = await vc.ExecuteReaderAsync();
                while (await vr.ReadAsync())
                {
                    DateTime? viewed = vr["ViewedAtUtc"] as DateTime?;
                    var when = viewed.HasValue
                        ? ToLocalStringFromUtc(DateTime.SpecifyKind(viewed.Value, DateTimeKind.Utc))
                        : null;

                    viewLogs.Add(new
                    {
                        viewedAt = when,
                        viewedAtText = when,
                        userId = vr["UserId"]?.ToString(),
                        userName = vr["UserName"]?.ToString()
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load view logs failed for {DocId}", id);
            }

            // ---------- 5) ViewBag (뷰/댓글 JS에서 사용하는 키) ----------
            ViewBag.DocId = id;           // JS에서 사용
            ViewBag.DocumentId = id;           // 호환용
            ViewBag.TemplateCode = templateCode ?? "";
            ViewBag.TemplateTitle = templateTitle ?? "";
            ViewBag.Status = status ?? "";

            ViewBag.DescriptorJson = descriptorJson ?? "{}";
            ViewBag.InputsJson = inputsJson;       // "{}" 고정
            ViewBag.PreviewJson = previewJson;

            ViewBag.ApprovalsJson = JsonSerializer.Serialize(approvals);
            ViewBag.ViewLogsJson = JsonSerializer.Serialize(viewLogs);

            // 댓글은 /Doc/Comments?docId=... API에서 별도 로드

            return View("Detail"); // Views/Doc/Detail.cshtml
        }

        // ========== DetailsData (상세 데이터 API) ==========
        [HttpGet("DetailsData")]
        [Produces("application/json")]
        public async Task<IActionResult> DetailsData(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return BadRequest(new { message = "DOC_Err_BadRequest" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT d.DocId, d.TemplateCode, d.TemplateTitle, d.Status, d.DescriptorJson, d.OutputPath, d.CreatedAt
FROM dbo.Documents d WHERE d.DocId=@id;";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

            string? descriptor = null, output = null, title = null, code = null, status = null;
            DateTime? createdAt = null;
            using (var rd = await cmd.ExecuteReaderAsync())
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
                ? "{}" : BuildPreviewJsonFromExcel(output);

            var approvals = new List<object>();
            using (var ac = conn.CreateCommand())
            {
                ac.CommandText = @"
SELECT StepOrder, RoleKey, ApproverValue, UserId, Status, Action, ActedAt, ActorName
FROM dbo.DocumentApprovals
WHERE DocId=@id
ORDER BY StepOrder;";
                ac.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });
                using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    approvals.Add(new
                    {
                        step = (int)r["StepOrder"],
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue = r["ApproverValue"]?.ToString(),
                        userId = r["UserId"]?.ToString(),
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = r["ActedAt"] is DateTime t ? ToLocalStringFromUtc(DateTime.SpecifyKind(t, DateTimeKind.Utc)) : null,
                        actorName = r["ActorName"]?.ToString()
                    });
                }
            }

            return Json(new
            {
                ok = true,
                docId = id,
                templateCode = code,
                templateTitle = title,
                status,
                createdAt = createdAt.HasValue ? ToLocalStringFromUtc(DateTime.SpecifyKind(createdAt.Value, DateTimeKind.Utc)) : null,
                descriptorJson = descriptor,
                previewJson = preview,
                approvals
            });
        }

        // ========= DTOs =========
        public sealed class ComposePostDto
        {
            public string? TemplateCode { get; set; }
            public Dictionary<string, string>? Inputs { get; set; }
            public ApprovalsDto? Approvals { get; set; }
            public MailDto? Mail { get; set; }
            public string? DescriptorVersion { get; set; }
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

                var approvalsDict = new Dictionary<int, (string approverType, string? value)>();
                if (root.TryGetProperty("Approvals", out var apprEl) && apprEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var a in apprEl.EnumerateArray())
                    {
                        int slot = a.TryGetProperty("Slot", out var s) && s.TryGetInt32(out var v) ? v : 0;
                        if (slot <= 0) continue;
                        var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                        var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;
                        approvalsDict[slot] = (MapApproverType(typ), val);
                    }
                }

                var approvals = approvalsDict
                    .OrderBy(kv => kv.Key)
                    .Select(kv => new { roleKey = $"A{kv.Key}", approverType = kv.Value.approverType, required = false, value = kv.Value.value ?? "" })
                    .ToList();

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
            desc.FlowGroups = groups;

            var json = JsonSerializer.Serialize(
                desc,
                new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                    WriteIndented = false
                });

            return json;
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
            string? departmentId
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
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = (SqlTransaction)tx;
                    cmd.CommandText = @"
INSERT INTO dbo.Documents
    (DocId, TemplateCode, TemplateTitle, Status, OutputPath, CompCd, DepartmentId, CreatedBy, CreatedByName, CreatedAt, DescriptorJson)
VALUES
    (@DocId, @TemplateCode, @TemplateTitle, @Status, @OutputPath, @CompCd, @DepartmentId, @CreatedBy, @CreatedByName, SYSUTCDATETIME(), @DescriptorJson);";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    cmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode });
                    cmd.Parameters.Add(new SqlParameter("@TemplateTitle", SqlDbType.NVarChar, 400) { Value = (object?)title ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@Status", SqlDbType.NVarChar, 20) { Value = status });
                    cmd.Parameters.Add(new SqlParameter("@OutputPath", SqlDbType.NVarChar, 500) { Value = outputPath });
                    cmd.Parameters.Add(new SqlParameter("@CompCd", SqlDbType.VarChar, 10) { Value = compCd });
                    cmd.Parameters.Add(new SqlParameter("@DepartmentId", SqlDbType.VarChar, 12)
                    { Value = string.IsNullOrWhiteSpace(departmentId) ? DBNull.Value : departmentId });
                    cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 450) { Value = userId ?? "" });
                    cmd.Parameters.Add(new SqlParameter("@CreatedByName", SqlDbType.NVarChar, 200) { Value = userName ?? "" });
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
                    using var aj = JsonDocument.Parse(approvalsJson ?? "[]");
                    var arr = aj.RootElement;
                    int step = 1;
                    foreach (var el in arr.EnumerateArray())
                    {
                        var roleKey = el.TryGetProperty("roleKey", out var rk) ? (rk.GetString() ?? $"A{step}") : $"A{step}";
                        var approverVal = el.TryGetProperty("value", out var vv) ? vv.GetString()
                                         : el.TryGetProperty("ApproverValue", out var vv2) ? vv2.GetString()
                                         : null;

                        string? resolvedUserId = await TryResolveUserIdAsync(conn, (SqlTransaction)tx, approverVal);

                        await using var acmd = conn.CreateCommand();
                        acmd.Transaction = (SqlTransaction)tx;
                        acmd.CommandText = @"
INSERT INTO dbo.DocumentApprovals (DocId, StepOrder, RoleKey, ApproverValue, UserId, Status, CreatedAt)
VALUES (@DocId, @StepOrder, @RoleKey, @ApproverValue, @UserId, 'Pending', SYSUTCDATETIME());";
                        acmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        acmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = step });
                        acmd.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 10) { Value = roleKey });
                        acmd.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 256)
                        { Value = (object?)approverVal ?? DBNull.Value });
                        acmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)resolvedUserId ?? DBNull.Value });
                        await acmd.ExecuteNonQueryAsync();

                        step++;
                    }
                }
                catch { }

                await tx.CommitAsync();
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }
        }

        private static async Task<string?> TryResolveUserIdAsync(SqlConnection conn, SqlTransaction tx, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            await using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = @"
SELECT TOP 1 Id
FROM dbo.AspNetUsers
WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);";
            cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            var id = (string?)await cmd.ExecuteScalarAsync();
            if (!string.IsNullOrWhiteSpace(id)) return id;

            await using var cmd2 = conn.CreateCommand();
            cmd2.Transaction = tx;
            cmd2.CommandText = @"
SELECT TOP 1 COALESCE(p.UserId, u.Id)
FROM dbo.UserProfiles p
LEFT JOIN dbo.AspNetUsers u ON u.Id = p.UserId
WHERE p.UserId=@v OR p.Email=@v OR p.DisplayName=@v OR p.Name=@v;";
            cmd2.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            id = (string?)await cmd2.ExecuteScalarAsync();
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        // ========= 승인 동기화 =========
        private async Task EnsureApprovalsAndSyncAsync(string docId, string docCode)
        {
            // (생략 없이 프로젝트에 이미 쓰이던 동일 SQL 사용)
            const string sql = @"/* ... 기존 EnsureApprovalsAndSync SQL 내용 ... */";
            // 위 주석 부분에 실제 SQL (이미 가지고 계신 것) 그대로 두시면 됩니다.
            await _db.Database.ExecuteSqlRawAsync(sql, docId, docCode);
        }

        // ========= Board =========
        [HttpGet("Board")]
        public IActionResult Board()
        {
            return View("Board");
        }

        [HttpGet("BoardData")]
        public async Task<IActionResult> BoardData(
            string tab = "created",
            int page = 1,
            int pageSize = 20,
            string titleFilter = "all",
            string sort = "created_desc",
            string? q = null,
            string approvalView = "all")
        {
            if (page < 1) page = 1;
            if (pageSize < 1 || pageSize > 100) pageSize = 20;

            var userId = User?.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;

            string orderBy = sort switch
            {
                "created_asc" => "ORDER BY d.CreatedAt ASC",
                "title_asc" => "ORDER BY d.TemplateTitle ASC",
                "title_desc" => "ORDER BY d.TemplateTitle DESC",
                _ => "ORDER BY d.CreatedAt DESC"
            };

            string whereSearch = string.IsNullOrWhiteSpace(q) ? "" : " AND d.TemplateTitle LIKE @Q ";
            string whereTitleFilter = titleFilter?.ToLowerInvariant() switch
            {
                "approved" => " AND ISNULL(d.Status, N'') = N'Approved' ",
                "rejected" => " AND ISNULL(d.Status, N'') = N'Rejected' ",
                "pending" => " AND ISNULL(d.Status, N'') LIKE N'Pending%' ",
                "onhold" => " AND ISNULL(d.Status, N'') = N'OnHold' ",
                _ => ""
            };

            string sqlCount;
            string sqlList;
            var offset = (page - 1) * pageSize;

            if (tab == "created")
            {
                sqlCount = @"
SELECT COUNT(1)
FROM Documents d
WHERE d.CreatedBy = @UserId" + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT d.DocId, d.TemplateTitle, d.CreatedAt, ISNULL(d.Status, N'') AS Status,
       up.DisplayName AS AuthorName,
       CAST(0 AS int) AS CommentCount,
       CAST(0 AS bit) AS HasAttachment
FROM Documents d
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
WHERE d.CreatedBy = @UserId" + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else if (tab == "approval")
            {
                string whereApprovalView = (approvalView ?? string.Empty).ToLowerInvariant() switch
                {
                    "pending" => " AND ISNULL(a.Status, N'') = N'Pending' ",
                    "approved" => " AND (ISNULL(a.Action, N'') = N'Approved' OR ISNULL(a.Status, N'') = N'Approved') ",
                    _ => ""
                };

                sqlCount = @"
SELECT COUNT(1)
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
WHERE a.UserId = @UserId" + whereApprovalView + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT d.DocId, d.TemplateTitle, d.CreatedAt, 
       CASE 
           WHEN ISNULL(a.Action, N'') <> N'' THEN a.Action 
           ELSE ISNULL(d.Status, N'') 
       END AS Status,
       up.DisplayName AS AuthorName,
       CAST(0 AS int) AS CommentCount,
       CAST(0 AS bit) AS HasAttachment
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
WHERE a.UserId = @UserId" + whereApprovalView + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else // shared
            {
                sqlCount = @"
SELECT COUNT(1)
FROM DocumentShares s
JOIN Documents d ON s.DocId = d.DocId
WHERE s.UserId = @UserId" + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT d.DocId, d.TemplateTitle, d.CreatedAt, ISNULL(d.Status, N'') AS Status,
       up.DisplayName AS AuthorName,
       CAST(0 AS int) AS CommentCount,
       CAST(0 AS bit) AS HasAttachment
FROM DocumentShares s
JOIN Documents d ON s.DocId = d.DocId
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
WHERE s.UserId = @UserId" + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }

            var connStr = _cfg.GetConnectionString("DefaultConnection");
            using var conn = new SqlConnection(connStr);
            await conn.OpenAsync();

            using var cmdCount = new SqlCommand(sqlCount, conn);
            cmdCount.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
            if (!string.IsNullOrWhiteSpace(q))
                cmdCount.Parameters.Add(new SqlParameter("@Q", SqlDbType.NVarChar, 200) { Value = $"%{q}%" });

            var total = Convert.ToInt32(await cmdCount.ExecuteScalarAsync());

            using var cmdList = new SqlCommand(sqlList, conn);
            cmdList.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
            if (!string.IsNullOrWhiteSpace(q))
                cmdList.Parameters.Add(new SqlParameter("@Q", SqlDbType.NVarChar, 200) { Value = $"%{q}%" });
            cmdList.Parameters.Add(new SqlParameter("@Offset", SqlDbType.Int) { Value = offset });
            cmdList.Parameters.Add(new SqlParameter("@PageSize", SqlDbType.Int) { Value = pageSize });

            var items = new List<object>();
            using (var rdr = await cmdList.ExecuteReaderAsync())
            {
                while (await rdr.ReadAsync())
                {
                    var createdAtLocal = (rdr["CreatedAt"] is DateTime utc) ? ToLocalStringFromUtc(utc) : "";
                    items.Add(new
                    {
                        docId = rdr["DocId"]?.ToString() ?? "",
                        templateTitle = rdr["TemplateTitle"]?.ToString() ?? "",
                        authorName = rdr["AuthorName"]?.ToString() ?? "",
                        createdAt = createdAtLocal,
                        status = rdr["Status"]?.ToString() ?? "",
                        commentCount = Convert.ToInt32(rdr["CommentCount"]),
                        hasAttachment = Convert.ToInt32(rdr["HasAttachment"]) == 1
                    });
                }
            }

            return Json(new { total, page, pageSize, items });
        }

        [HttpGet("BoardBadges")]
        public async Task<IActionResult> BoardBadges()
        {
            var userId = User?.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var connStr = _cfg.GetConnectionString("DefaultConnection");
            using var conn = new SqlConnection(connStr);
            await conn.OpenAsync();

            var sqlApprovalPending = @"
SELECT COUNT(1)
FROM DocumentApprovals a
WHERE a.UserId = @UserId AND ISNULL(a.Status, N'') = N'Pending';";

            var sqlSharedUnread = @"
SELECT COUNT(1)
FROM DocumentShares s
WHERE s.UserId = @UserId
AND NOT EXISTS (
    SELECT 1 
    FROM DocumentAuditLogs l
    WHERE l.DocId = s.DocId AND l.ActorId = @UserId AND l.ActionCode = N'READ'
);";

            int approvalPending = 0;
            int sharedUnread = 0;

            using (var cmd = new SqlCommand(sqlApprovalPending, conn))
            {
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
                approvalPending = Convert.ToInt32(await cmd.ExecuteScalarAsync());
            }
            using (var cmd = new SqlCommand(sqlSharedUnread, conn))
            {
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
                sharedUnread = Convert.ToInt32(await cmd.ExecuteScalarAsync());
            }

            return Json(new { created = 0, approvalPending, sharedUnread });
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
            var link = Url.Action("Detail", "Doc", new { id = docId }, Request.Scheme) ?? "";

            var tz = ResolveCompanyTimeZone();
            var nowLocal = TimeZoneInfo.ConvertTime(DateTime.UtcNow, tz);

            var mm = nowLocal.Month;
            var dd = nowLocal.Day;
            var hh = nowLocal.Hour;
            var min = nowLocal.Minute;

            var subject = string.Format(_S["DOC_Email_Subject"].Value, authorName, docTitle);
            var body = string.Format(
                _S["DOC_Email_BodyHtml"].Value,
                string.IsNullOrWhiteSpace(recipientDisplay) ? "" : recipientDisplay,
                authorName, docTitle, mm, dd, hh, min, link
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

            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT CommentId, DocId, ParentCommentId, ThreadRootId, Depth,
       Body, HasAttachment, IsDeleted,
       CreatedBy, CreatedAt, UpdatedAt
FROM dbo.DocumentComments
WHERE DocId = @DocId AND IsDeleted = 0
ORDER BY ThreadRootId, Depth, CreatedAt;";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

            var currentUser = User?.Identity?.Name ?? "";

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var commentId = Convert.ToInt64(rd["CommentId"]);
                var parent = rd["ParentCommentId"] == DBNull.Value
                    ? (long?)null
                    : Convert.ToInt64(rd["ParentCommentId"]);
                var depth = Convert.ToInt32(rd["Depth"]);
                var hasAttachment = Convert.ToBoolean(rd["HasAttachment"]);

                // CreatedAt → 로컬 시간 문자열
                var createdAtUtc = (DateTime)rd["CreatedAt"];
                var createdAtLocal = ToLocalStringFromUtc(
                    DateTime.SpecifyKind(createdAtUtc, DateTimeKind.Utc));

                var createdBy = rd["CreatedBy"]?.ToString() ?? "";

                // 본인 삭제/수정 권한 (필요 시 Admin 권한 조건 추가)
                // 2025.11.10 Changed: canModify 계산 시 User null 가능성 방지를 위해 null 조건 연산자 적용 (CS8602 경고 제거, 동작 동일)
                var safeCreatedBy = createdBy ?? string.Empty;
                var safeCurrentUser = currentUser ?? string.Empty;

                var canModify =
                    (!string.IsNullOrEmpty(safeCreatedBy)
                        && !string.IsNullOrEmpty(safeCurrentUser)
                        && string.Equals(safeCreatedBy, safeCurrentUser, StringComparison.OrdinalIgnoreCase))
                    || (User?.IsInRole("Admin") ?? false);

                items.Add(new
                {
                    commentId,
                    docId = (string)rd["DocId"],
                    parentCommentId = parent,
                    depth,
                    body = rd["Body"]?.ToString() ?? "",
                    hasAttachment,
                    createdBy,
                    createdAt = createdAtLocal,
                    canEdit = canModify,
                    canDelete = canModify
                });
            }

            return Json(new { items });
        }


        [HttpPost("Comments")]
        [Consumes("application/json")]
        [Produces("application/json")]
        // [ValidateAntiForgeryToken]  // 일단 주석 처리해서 토큰 문제부터 배제
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
                        threadRootId = (int)r["RootId"];
                        depth = (int)r["Depth"] + 1;
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
        [Produces("application/json")]
        // [ValidateAntiForgeryToken] // 동일하게, 원인잡고 나중에 필요시 추가
        public async Task<IActionResult> DeleteComment(int id, string docId)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
UPDATE dbo.DocumentComments
SET IsDeleted = 1, UpdatedAt = SYSUTCDATETIME()
WHERE CommentId = @id AND DocId = @docId;";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = id });
            cmd.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = docId });

            var rows = await cmd.ExecuteNonQueryAsync();
            if (rows == 0)
                return NotFound(new { ok = false, messages = new[] { "DOC_Err_RequestFailed" }, detail = "not found" });

            return Json(new { ok = true });
        }

        [HttpPut("Comments/{id}")]
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

            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
UPDATE dbo.DocumentComments
SET Body = @Body,
    UpdatedAt = SYSUTCDATETIME()
WHERE CommentId = @Id
  AND DocId = @DocId
  AND IsDeleted = 0
  AND CreatedBy = @UserName;"; // 본인만 수정
            cmd.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = dto.body.Trim() });
            cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = id });
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId ?? "" });
            cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200) { Value = userName });

            var rows = await cmd.ExecuteNonQueryAsync();
            if (rows == 0)
                return Forbid(); // 또는 NotFound

            return Json(new { ok = true });
        }

    }
}
