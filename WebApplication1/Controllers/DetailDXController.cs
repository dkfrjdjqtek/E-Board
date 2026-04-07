// 2026.03.17 Added: DetailDX 구성
// 2026.03.18 Changed: DxTemp 임시파일 생성 제거 — ReadOnly 조회이므로 원본 파일 직접 사용
// 2026.03.23 Added: Cooperate 기능 추가, DetailController 전체 기능 마이그레이션
// 2026.03.24 Added: 협조 스탬프 지원 — DescriptorJson.Cooperations에서 CellA1 파싱, UserProfiles.SignatureRelativePath JOIN
// 2026.03.25 Fixed: 협조 셀맵 파싱 시 cellA1 키 추가 탐색 (Cell.A1 / A1 / cellA1 / CellA1 순서로 탐색)
// 2026.03.26 Fixed: Documents.DescriptorJson에 cellA1 없을 때 DocTemplateVersion.DescriptorJson Cooperations[].Cell.A1 fallback 추가
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WebApplication1.Models;
using WebApplication1.Services;
using ClosedXML.Excel;
using static WebApplication1.Controllers.DocControllerHelper;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class DetailDXController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IConfiguration _cfg;
        private readonly IWebHostEnvironment _env;
        private readonly ILogger<DetailDXController> _log;
        private readonly IWebPushNotifier _webPushNotifier;
        private readonly DocControllerHelper _helper;

        public DetailDXController(
            IStringLocalizer<SharedResource> S,
            IConfiguration cfg,
            IWebHostEnvironment env,
            ILogger<DetailDXController> log,
            IWebPushNotifier webPushNotifier)
        {
            _S = S;
            _cfg = cfg;
            _env = env;
            _log = log;
            _webPushNotifier = webPushNotifier;
            _helper = new DocControllerHelper(_cfg, _env, () => User);
        }

        // ========== DetailDX (GET) ==========
        [HttpGet("DetailDX/{id?}")]
        public async Task<IActionResult> DetailDX(string? id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                var qid = HttpContext?.Request?.Query["id"].ToString();
                if (!string.IsNullOrWhiteSpace(qid))
                    id = qid;
            }

            if (string.IsNullOrWhiteSpace(id))
                return Redirect("/Doc/Board");

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            string? templateCode = null;
            string? templateTitle = null;
            string? status = null;
            string? descriptorJson = null;
            string? outputPath = null;
            string? compCd = null;
            string? creatorUserId = null;
            string? creatorNameRaw = null;
            DateTime? createdAtUtc = null;

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
       COALESCE(p.DisplayName, d.CreatedByName, u.UserName, u.Email) AS CreatorDisplayName
FROM dbo.Documents d
LEFT JOIN dbo.AspNetUsers  u ON u.Id     = d.CreatedBy
LEFT JOIN dbo.UserProfiles p ON p.UserId = d.CreatedBy
WHERE d.DocId = @DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var rd = await cmd.ExecuteReaderAsync();
                if (!await rd.ReadAsync())
                {
                    ViewBag.Message = _S["DOC_Err_DocumentNotFound"];
                    ViewBag.DocId = id;
                    ViewBag.DocumentId = id;
                    return View("~/Views/Doc/DetailDX.cshtml");
                }

                templateCode = rd["TemplateCode"] as string ?? string.Empty;
                templateTitle = rd["TemplateTitle"] as string ?? string.Empty;
                status = rd["Status"] as string ?? string.Empty;
                descriptorJson = rd["DescriptorJson"] as string ?? "{}";
                outputPath = rd["OutputPath"] as string ?? string.Empty;
                compCd = rd["CompCd"] as string ?? string.Empty;
                creatorUserId = rd["CreatedBy"] as string ?? string.Empty;
                creatorNameRaw = rd["CreatorDisplayName"] as string ?? string.Empty;
                createdAtUtc = rd["CreatedAt"] is DateTime cdt ? cdt : (DateTime?)null;
            }

            string? creatorDisplayName = creatorNameRaw;
            if (!string.IsNullOrWhiteSpace(creatorUserId))
            {
                await using var pc = conn.CreateCommand();
                pc.CommandText = @"
SELECT TOP 1 DisplayName
FROM dbo.UserProfiles
WHERE UserId = @UserId;";
                pc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = creatorUserId! });
                var objName = await pc.ExecuteScalarAsync();
                if (objName != null && objName != DBNull.Value)
                    creatorDisplayName = Convert.ToString(objName) ?? creatorDisplayName;
            }

            string createdAtText = string.Empty;
            if (createdAtUtc.HasValue)
                createdAtText = _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(createdAtUtc.Value, DateTimeKind.Utc));

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
                rcCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });
                var obj = await rcCmd.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                {
                    var dt = (obj is DateTime dtx) ? dtx : Convert.ToDateTime(obj);
                    recalledAtText = _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(dt, DateTimeKind.Utc));
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: recalled time load failed for {DocId}", id);
            }

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
                tvCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 50) { Value = templateCode! });
                var obj = await tvCmd.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                    templateVersionId = Convert.ToInt32(obj);
            }

            var dxOpenAbsPath = ResolveContentRootRelativePath(outputPath) ?? string.Empty;
            _log.LogInformation("DetailDX: dxOpenAbsPath={Path} exists={Exists}", dxOpenAbsPath, System.IO.File.Exists(dxOpenAbsPath ?? ""));
            var previewJson = "{}";
            try
            {
                if (!string.IsNullOrWhiteSpace(dxOpenAbsPath) && System.IO.File.Exists(dxOpenAbsPath))
                    previewJson = BuildPreviewJsonFromExcel(dxOpenAbsPath, maxRows: 500, maxCols: 100);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: preview json build failed for {DocId}", id);
                previewJson = "{}";
            }

            // ★ 추가: previewJson={}이면 템플릿 previewJson fallback
            if ((previewJson == "{}" || string.IsNullOrWhiteSpace(previewJson))
                && !string.IsNullOrWhiteSpace(templateCode))
            {
                try
                {
                    await using var tvFbCmd = conn.CreateCommand();
                    tvFbCmd.CommandText = @"
SELECT TOP 1 v.PreviewJson
FROM dbo.DocTemplateVersion v
JOIN dbo.DocTemplateMaster m ON m.Id = v.TemplateId
WHERE m.DocCode = @TemplateCode
ORDER BY v.VersionNo DESC, v.Id DESC;";
                    tvFbCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode! });
                    var tplPreview = await tvFbCmd.ExecuteScalarAsync() as string;
                    if (!string.IsNullOrWhiteSpace(tplPreview) && tplPreview != "{}")
                        previewJson = tplPreview;
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "DetailDX: template previewJson fallback failed for {DocId}", id);
                }
            }

            var openRel = MakeRelativeToContentRoot(dxOpenAbsPath);

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
LEFT JOIN dbo.UserProfiles up ON up.UserId = a.UserId
WHERE a.DocId = @DocId
ORDER BY a.StepOrder;";
                ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    DateTime? acted = r.IsDBNull(6) ? (DateTime?)null : r.GetDateTime(6);
                    approvals.Add(new
                    {
                        step = r.GetInt32(0),
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue = r["ApproverValue"]?.ToString(),
                        userId = r["UserId"]?.ToString(),
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = acted.HasValue
                            ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                            : string.Empty,
                        actorName = r["ActorName"]?.ToString(),
                        approverDisplayText = r["ApproverDisplayText"]?.ToString() ?? string.Empty,
                        signaturePath = r["SignaturePath"]?.ToString()
                    });
                }
            }

            var approvalCells = new List<object>();
            if (templateVersionId.HasValue)
            {
                await using var ac2 = conn.CreateCommand();
                ac2.CommandText = @"
SELECT Slot, ISNULL(CellA1, A1) AS A1
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

            var docFiles = new List<object>();
            try
            {
                await using var fc = conn.CreateCommand();
                fc.CommandText = @"
SELECT FileKey, OriginalName, ContentType, ByteSize, UploadedAt
FROM dbo.DocumentFiles
WHERE DocId = @DocId
ORDER BY UploadedAt, FileId ASC;";
                fc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var fr = await fc.ExecuteReaderAsync();
                while (await fr.ReadAsync())
                {
                    var uploadedAtText = string.Empty;
                    if (fr["UploadedAt"] != DBNull.Value)
                    {
                        var upUtc = Convert.ToDateTime(fr["UploadedAt"]);
                        uploadedAtText = _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(upUtc, DateTimeKind.Utc));
                    }
                    docFiles.Add(new
                    {
                        FileKey = fr["FileKey"] as string ?? string.Empty,
                        OriginalName = fr["OriginalName"] as string ?? string.Empty,
                        ContentType = fr["ContentType"] as string ?? string.Empty,
                        ByteSize = fr["ByteSize"] != DBNull.Value ? Convert.ToInt64(fr["ByteSize"]) : 0L,
                        uploadedAtText
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load DocumentFiles failed for {DocId}", id);
            }

            // ========== 협조 셀 맵 ==========
            // 1순위: Documents.DescriptorJson cooperations[].cellA1
            // 2순위: DocTemplateVersion.DescriptorJson Cooperations[].Cell.A1  ← fallback
            var cooperationCellMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var cooperationRoleCellMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // ── 1순위: Documents.DescriptorJson ──────────────────────────────
            try
            {
                if (!string.IsNullOrWhiteSpace(descriptorJson) && descriptorJson != "{}")
                {
                    using var descDoc = JsonDocument.Parse(descriptorJson);

                    JsonElement coopsEl;
                    var hasCoops =
                        (descDoc.RootElement.TryGetProperty("Cooperations", out coopsEl)
                      || descDoc.RootElement.TryGetProperty("cooperations", out coopsEl))
                        && coopsEl.ValueKind == JsonValueKind.Array;

                    if (hasCoops)
                    {
                        foreach (var coop in coopsEl.EnumerateArray())
                        {
                            var av =
                                TryGetJsonString(coop, "ApproverValue")
                                ?? TryGetJsonString(coop, "approverValue")
                                ?? TryGetJsonString(coop, "value")
                                ?? string.Empty;

                            var roleKey =
                                TryGetJsonString(coop, "RoleKey")
                                ?? TryGetJsonString(coop, "roleKey")
                                ?? string.Empty;

                            var a1 = string.Empty;

                            if (coop.TryGetProperty("Cell", out var cellEl) && cellEl.ValueKind == JsonValueKind.Object)
                                a1 = TryGetJsonString(cellEl, "A1") ?? TryGetJsonString(cellEl, "a1") ?? string.Empty;

                            if (string.IsNullOrWhiteSpace(a1) &&
                                coop.TryGetProperty("cell", out var cellEl2) && cellEl2.ValueKind == JsonValueKind.Object)
                                a1 = TryGetJsonString(cellEl2, "A1") ?? TryGetJsonString(cellEl2, "a1") ?? string.Empty;

                            if (string.IsNullOrWhiteSpace(a1))
                                a1 = TryGetJsonString(coop, "A1") ?? TryGetJsonString(coop, "a1") ?? string.Empty;

                            if (string.IsNullOrWhiteSpace(a1))
                                a1 = TryGetJsonString(coop, "cellA1") ?? TryGetJsonString(coop, "CellA1") ?? string.Empty;

                            if (!string.IsNullOrWhiteSpace(av) && !string.IsNullOrWhiteSpace(a1))
                                cooperationCellMap[av] = a1;

                            if (!string.IsNullOrWhiteSpace(roleKey) && !string.IsNullOrWhiteSpace(a1))
                                cooperationRoleCellMap[roleKey] = a1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: DescriptorJson Cooperations parse failed for {DocId}", id);
            }

            // ── 2순위 fallback: DocTemplateVersion.DescriptorJson Cooperations[].Cell.A1 ──
            // Documents.DescriptorJson에 cellA1이 없는 기존 문서를 위한 보완
            if (cooperationRoleCellMap.Count == 0 && !string.IsNullOrWhiteSpace(templateCode))
            {
                try
                {
                    await using var tvFbCmd = conn.CreateCommand();
                    tvFbCmd.CommandText = @"
SELECT TOP 1 v.DescriptorJson
FROM dbo.DocTemplateVersion v
JOIN dbo.DocTemplateMaster m ON m.Id = v.TemplateId
WHERE m.DocCode = @TemplateCode
ORDER BY v.VersionNo DESC, v.Id DESC;";
                    tvFbCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode! });

                    var rawDesc = await tvFbCmd.ExecuteScalarAsync() as string;
                    if (!string.IsNullOrWhiteSpace(rawDesc))
                    {
                        using var rawDoc = JsonDocument.Parse(rawDesc);
                        JsonElement coopsEl2;
                        var hasCoops2 =
                            (rawDoc.RootElement.TryGetProperty("Cooperations", out coopsEl2)
                          || rawDoc.RootElement.TryGetProperty("cooperations", out coopsEl2))
                            && coopsEl2.ValueKind == JsonValueKind.Array;

                        if (hasCoops2)
                        {
                            int seq = 1;
                            foreach (var coop in coopsEl2.EnumerateArray())
                            {
                                var roleKey = $"C{seq}";

                                // ApproverValue (GUID) — 맵 키로만 사용
                                var av =
                                    TryGetJsonString(coop, "ApproverValue")
                                    ?? TryGetJsonString(coop, "approverValue")
                                    ?? TryGetJsonString(coop, "value")
                                    ?? string.Empty;

                                // A1: Cell.A1 우선
                                var a1 = string.Empty;
                                if (coop.TryGetProperty("Cell", out var cellEl) && cellEl.ValueKind == JsonValueKind.Object)
                                    a1 = TryGetJsonString(cellEl, "A1") ?? TryGetJsonString(cellEl, "a1") ?? string.Empty;

                                if (string.IsNullOrWhiteSpace(a1))
                                    a1 = TryGetJsonString(coop, "A1") ?? TryGetJsonString(coop, "a1") ?? string.Empty;

                                if (string.IsNullOrWhiteSpace(a1))
                                    a1 = TryGetJsonString(coop, "cellA1") ?? TryGetJsonString(coop, "CellA1") ?? string.Empty;

                                if (!string.IsNullOrWhiteSpace(a1))
                                {
                                    cooperationRoleCellMap[roleKey] = a1;
                                    if (!string.IsNullOrWhiteSpace(av))
                                        cooperationCellMap[av] = a1;
                                }

                                seq++;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "DetailDX: template DescriptorJson Cooperations cellA1 fallback failed for {DocId}", id);
                }
            }

            // ========== 협조 데이터 로드 ==========
            var cooperations = new List<object>();
            var approvalTableRows = new List<ApprovalTableRowVm>
            {
                new ApprovalTableRowVm
                {
                    RowType     = "drafter",
                    GroupOrder  = 0,
                    SortOrder   = 0,
                    RoleKey     = "DRAFT",
                    RoleText    = _S["DOC_Role_Submit"],
                    DisplayName = creatorDisplayName ?? string.Empty,
                    Status      = string.Equals(status, "Recalled", StringComparison.OrdinalIgnoreCase) ? "Recalled" : "Created",
                    ActedAtText = string.Equals(status, "Recalled", StringComparison.OrdinalIgnoreCase) ? recalledAtText : createdAtText
                }
            };

            foreach (var a in approvals)
            {
                var roleKey = GetAnonymousString(a, "roleKey");
                var approverDisplayText = GetAnonymousString(a, "approverDisplayText");
                var approverValue = GetAnonymousString(a, "approverValue");

                approvalTableRows.Add(new ApprovalTableRowVm
                {
                    RowType = "approval",
                    GroupOrder = 1,
                    SortOrder = GetAnonymousInt(a, "step") ?? ParseRoleOrder(roleKey),
                    RoleKey = roleKey,
                    RoleText = BuildRoleText(roleKey, "approval"),
                    DisplayName = !string.IsNullOrWhiteSpace(approverDisplayText) ? approverDisplayText : approverValue,
                    Status = GetAnonymousString(a, "status"),
                    ActedAtText = GetAnonymousString(a, "actedAtText")
                });
            }

            try
            {
                await using var cc = conn.CreateCommand();
                cc.CommandText = @"
SELECT dc.Id,
       dc.RoleKey,
       dc.UserId,
       dc.ApproverValue,
       dc.Status,
       dc.Action,
       dc.ActedAt,
       dc.ActorName,
       COALESCE(dc.ApproverDisplayText, up.DisplayName, dc.ActorName, dc.ApproverValue) AS ApproverDisplayText,
       up.SignatureRelativePath AS SignaturePath
FROM dbo.DocumentCooperations dc
LEFT JOIN dbo.UserProfiles up ON up.UserId = dc.UserId
WHERE dc.DocId = @DocId
ORDER BY dc.RoleKey;";
                cc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var cr = await cc.ExecuteReaderAsync();
                while (await cr.ReadAsync())
                {
                    DateTime? acted = cr["ActedAt"] is DateTime dt ? dt : (DateTime?)null;
                    var roleKey = cr["RoleKey"]?.ToString() ?? string.Empty;
                    var uid = cr["UserId"]?.ToString() ?? string.Empty;
                    var av = cr["ApproverValue"]?.ToString() ?? string.Empty;
                    var approverDisplayText = cr["ApproverDisplayText"]?.ToString() ?? string.Empty;

                    // 셀 위치 결정: roleKey → userId → approverValue 순으로 맵 탐색
                    var cellA1 = string.Empty;
                    if (!string.IsNullOrWhiteSpace(roleKey) && cooperationRoleCellMap.TryGetValue(roleKey, out var byRole))
                        cellA1 = byRole;
                    else if (!string.IsNullOrWhiteSpace(uid) && cooperationCellMap.TryGetValue(uid, out var byUid))
                        cellA1 = byUid;
                    else if (!string.IsNullOrWhiteSpace(av) && cooperationCellMap.TryGetValue(av, out var byAv))
                        cellA1 = byAv;

                    cooperations.Add(new
                    {
                        roleKey,
                        userId = uid,
                        approverValue = av,
                        status = cr["Status"]?.ToString(),
                        action = cr["Action"]?.ToString(),
                        actedAtText = acted.HasValue
                            ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                            : string.Empty,
                        actorName = cr["ActorName"]?.ToString(),
                        approverDisplayText,
                        cellA1,
                        signaturePath = cr["SignaturePath"]?.ToString() ?? string.Empty
                    });

                    approvalTableRows.Add(new ApprovalTableRowVm
                    {
                        RowType = "cooperation",
                        GroupOrder = 2,
                        SortOrder = ParseRoleOrder(roleKey),
                        RoleKey = roleKey,
                        RoleText = BuildRoleText(roleKey, "cooperation"),
                        DisplayName = !string.IsNullOrWhiteSpace(approverDisplayText) ? approverDisplayText : av,
                        Status = cr["Status"]?.ToString() ?? string.Empty,
                        ActedAtText = acted.HasValue
                            ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                            : string.Empty
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load DocumentCooperations failed for {DocId}", id);
            }

            approvalTableRows = approvalTableRows
                .OrderBy(x => x.GroupOrder)
                .ThenBy(x => x.SortOrder)
                .ThenBy(x => x.RoleKey)
                .ToList();

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
                sc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var sr = await sc.ExecuteReaderAsync();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                while (await sr.ReadAsync())
                {
                    var uid = sr.IsDBNull(0) ? null : sr.GetString(0);
                    if (string.IsNullOrWhiteSpace(uid)) continue;
                    if (seen.Add(uid)) sharedRecipientUserIds.Add(uid);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load DocumentShares failed for {DocId}", id);
            }

            try
            {
                ViewBag.OrgTreeNodes = await _helper.BuildOrgTreeNodesAsync("ko");
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load OrgTreeNodes failed for {DocId}", id);
                ViewBag.OrgTreeNodes = Array.Empty<OrgTreeNode>();
            }

            try
            {
                await LogDocumentViewAsync(conn, id!);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: LogDocumentViewAsync failed for {DocId}", id);
            }

            var viewLogs = new List<object>();
            try
            {
                await using var vc = conn.CreateCommand();
                vc.CommandText = @"
SELECT TOP 200
       v.ViewedAt, v.ViewerId, v.ViewerRole, v.ClientIp, v.UserAgent,
       COALESCE(p.DisplayName, u.UserName, u.Email) AS UserName
FROM dbo.DocumentViewLogs v
LEFT JOIN dbo.UserProfiles p ON p.UserId = v.ViewerId
LEFT JOIN dbo.AspNetUsers  u ON u.Id     = v.ViewerId
WHERE v.DocId = @DocId
ORDER BY v.ViewedAt DESC;";
                vc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var vr = await vc.ExecuteReaderAsync();
                while (await vr.ReadAsync())
                {
                    DateTime? viewedUtc = vr.IsDBNull(0) ? (DateTime?)null : vr.GetDateTime(0);
                    var viewedAtText = viewedUtc.HasValue
                        ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(viewedUtc.Value, DateTimeKind.Utc))
                        : string.Empty;
                    viewLogs.Add(new
                    {
                        viewedAt = viewedUtc,
                        viewedAtText,
                        userId = vr.IsDBNull(1) ? null : Convert.ToString(vr.GetValue(1)),
                        viewerRole = vr.IsDBNull(2) ? null : Convert.ToString(vr.GetValue(2)),
                        clientIp = vr.IsDBNull(3) ? null : Convert.ToString(vr.GetValue(3)),
                        userAgent = vr.IsDBNull(4) ? null : Convert.ToString(vr.GetValue(4)),
                        userName = vr.IsDBNull(5) ? null : Convert.ToString(vr.GetValue(5))
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load view logs failed for {DocId}", id);
            }

            string stampRectsJson = "{\"widthPx\":0,\"heightPx\":0,\"stamps\":[]}";
            try
            {
                if (!string.IsNullOrWhiteSpace(dxOpenAbsPath) && System.IO.File.Exists(dxOpenAbsPath))
                {
                    stampRectsJson = BuildStampRectsJsonFromExcel(
                        dxOpenAbsPath,
                        JsonSerializer.Serialize(approvalCells),
                        JsonSerializer.Serialize(approvals),
                        JsonSerializer.Serialize(cooperations),
                        _log);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: stamp rects build failed for {DocId}", id);
                stampRectsJson = "{\"widthPx\":0,\"heightPx\":0,\"stamps\":[]}";
            }

            var renderArea = new ExcelRenderAreaInfo
            {
                UsedMaxRow = 1,
                UsedMaxCol = 1,
                RenderMaxRow = 2,
                RenderMaxCol = 2,
                UsedLastCellA1 = "A1",
                RenderLastCellA1 = "B2",
                WidthPx = 0d,
                HeightPx = 0d
            };

            try
            {
                if (!string.IsNullOrWhiteSpace(dxOpenAbsPath) && System.IO.File.Exists(dxOpenAbsPath))
                    renderArea = GetRenderAreaInfoFromExcel(dxOpenAbsPath, _log);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: render area build failed for {DocId}", id);
            }

            string renderDiagJson = "{}";
            try
            {
                if (!string.IsNullOrWhiteSpace(dxOpenAbsPath) && System.IO.File.Exists(dxOpenAbsPath))
                {
                    renderDiagJson = BuildRenderAreaDiagJsonFromExcel(dxOpenAbsPath, _log);
                    _log.LogInformation("DetailDX RenderDiag {DocId} {Diag}", id, renderDiagJson);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: render diag build failed for {DocId}", id);
                renderDiagJson = "{}";
            }

            ViewBag.RenderDiagJson = renderDiagJson;


            ViewBag.RenderUsedLastCellA1 = renderArea.UsedLastCellA1;
            ViewBag.RenderLastCellA1 = renderArea.RenderLastCellA1;
            ViewBag.RenderMaxRow = renderArea.RenderMaxRow;
            ViewBag.RenderMaxCol = renderArea.RenderMaxCol;
            ViewBag.RenderWidthPx = Math.Ceiling(renderArea.WidthPx);
            ViewBag.RenderHeightPx = Math.Ceiling(renderArea.HeightPx);


            // 기존 ViewBag 할당부에 아래 한 줄 추가
            ViewBag.StampRectsJson = stampRectsJson;

            var caps = await GetApprovalCapabilitiesAsync(id!);

            ViewData["Title"] = _S["DOC_Title_Detail"];
            ViewData["DisableDxAll"] = true;

            ViewBag.DocId = id;
            ViewBag.DocumentId = id;
            ViewBag.TemplateCode = templateCode ?? string.Empty;
            ViewBag.TemplateTitle = templateTitle ?? string.Empty;
            ViewBag.Status = status ?? string.Empty;
            ViewBag.CompCd = compCd ?? string.Empty;
            ViewBag.CreatorName = creatorDisplayName ?? string.Empty;
            ViewBag.DescriptorJson = descriptorJson ?? "{}";
            ViewBag.InputsJson = "{}";
            ViewBag.PreviewJson = previewJson;
            ViewBag.ApprovalsJson = JsonSerializer.Serialize(approvals);
            ViewBag.ViewLogsJson = JsonSerializer.Serialize(viewLogs);
            ViewBag.ApprovalCellsJson = JsonSerializer.Serialize(approvalCells);
            ViewBag.DocumentFilesJson = JsonSerializer.Serialize(docFiles);
            ViewBag.CooperationsJson = JsonSerializer.Serialize(cooperations);
            ViewBag.ApprovalTableRowsJson = JsonSerializer.Serialize(approvalTableRows);
            ViewBag.SelectedRecipientUserIds = sharedRecipientUserIds;
            ViewBag.SelectedRecipientUserIdsJson = JsonSerializer.Serialize(sharedRecipientUserIds);
            ViewBag.CreatedAtText = createdAtText;
            ViewBag.RecalledAtText = recalledAtText;
            ViewBag.ExcelPath = dxOpenAbsPath ?? string.Empty;
            ViewBag.OpenRel = openRel ?? string.Empty;
            var request = HttpContext.Request;
            ViewBag.DxCallbackUrl = $"{request.Scheme}://{request.Host}/Doc/dx-callback";
            ViewBag.DxDocumentId = "detail_" + (id ?? Guid.NewGuid().ToString("N"));
            ViewBag.CanRecall = caps.canRecall;
            ViewBag.CanApprove = caps.canApprove;
            ViewBag.CanHold = caps.canHold;
            ViewBag.CanReject = caps.canReject;
            ViewBag.CanBackToNew = caps.canRecall;
            ViewBag.CanPrint = true;
            ViewBag.CanExportPdf = true;
            ViewBag.CanCooperate = caps.canCooperate;
            ViewBag.CanCooperateReject = caps.canCooperate;

            return View("~/Views/Doc/DetailDX.cshtml");
        }

        private static ExcelRenderAreaInfo GetRenderAreaInfoFromExcel(string excelPath, ILogger? log = null)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
                return new ExcelRenderAreaInfo();

            try
            {
                using var wb = new XLWorkbook(excelPath);
                var ws = wb.Worksheets.First();

                double defaultRowPt = ws.RowHeight <= 0 ? 15 : ws.RowHeight;
                double defaultColChar = ws.ColumnWidth <= 0 ? 8.43 : ws.ColumnWidth;

                var bbox = BuildVisibleBBox(ws);

                var usedMaxRow = Math.Max(1, bbox.usedMaxRow);
                var usedMaxCol = Math.Max(1, bbox.usedMaxCol);

                var renderMaxRow = Math.Min(XLHelper.MaxRowNumber, usedMaxRow + 1);
                var renderMaxCol = Math.Min(XLHelper.MaxColumnNumber, usedMaxCol + 1);

                double widthPx = 0d;
                for (int c = 1; c <= renderMaxCol; c++)
                {
                    var colWidth = ws.Column(c).Width <= 0 ? defaultColChar : ws.Column(c).Width;
                    widthPx += ExcelColWidthToPx(colWidth);
                }

                double heightPx = 0d;
                for (int r = 1; r <= renderMaxRow; r++)
                {
                    var rowHeight = ws.Row(r).Height <= 0 ? defaultRowPt : ws.Row(r).Height;
                    heightPx += RowHeightPtToPx(rowHeight);
                }

                return new ExcelRenderAreaInfo
                {
                    UsedMaxRow = usedMaxRow,
                    UsedMaxCol = usedMaxCol,
                    RenderMaxRow = renderMaxRow,
                    RenderMaxCol = renderMaxCol,
                    UsedLastCellA1 = ToA1(usedMaxRow, usedMaxCol),
                    RenderLastCellA1 = ToA1(renderMaxRow, renderMaxCol),
                    WidthPx = Math.Round(widthPx, 2),
                    HeightPx = Math.Round(heightPx, 2)
                };
            }
            catch (Exception ex)
            {
                log?.LogWarning(ex, "GetRenderAreaInfoFromExcel failed path={Path}", excelPath);
                return new ExcelRenderAreaInfo();
            }
        }

        private static double ExcelColWidthToPx(double widthChars)
        {
            const double mdw = 7d;
            var w = widthChars < 0 ? 0 : widthChars;
            return Math.Floor((w * mdw + 5d) / mdw) * mdw + 2d;
        }

        private static double RowHeightPtToPx(double heightPt)
        {
            return Math.Round((heightPt <= 0 ? 15d : heightPt) * 96d / 72d, 2);
        }


        private static bool HasVisibleBorder(IXLCell cell)
        {
            var b = cell.Style.Border;
            return b.LeftBorder != XLBorderStyleValues.None
                || b.RightBorder != XLBorderStyleValues.None
                || b.TopBorder != XLBorderStyleValues.None
                || b.BottomBorder != XLBorderStyleValues.None;
        }

        private static bool HasVisibleFill(IXLCell cell)
        {
            var f = cell.Style.Fill;
            if (f.PatternType != XLFillPatternValues.None) return true;

            try
            {
                var bg = f.BackgroundColor;
                if (bg != null && !bg.Equals(XLColor.NoColor) && !bg.Equals(XLColor.FromIndex(64)))
                    return true;
            }
            catch { }

            return false;
        }

        private static bool HasMeaningfulValue(IXLCell cell)
        {
            if (cell == null)
                return false;

            try
            {
                if (cell.DataType == XLDataType.Blank)
                    return false;

                var s = cell.GetString();
                if (!string.IsNullOrWhiteSpace(s))
                    return true;

                return cell.DataType == XLDataType.Number
                    || cell.DataType == XLDataType.DateTime
                    || cell.DataType == XLDataType.Boolean
                    || cell.DataType == XLDataType.TimeSpan
                    || cell.DataType == XLDataType.Error;
            }
            catch
            {
                try
                {
                    var fallback = cell.Value.ToString();
                    return !string.IsNullOrWhiteSpace(fallback);
                }
                catch
                {
                    return false;
                }
            }
        }

        private static bool TryGetMergeBounds(
            List<(int r1, int c1, int r2, int c2)> mergedRanges,
            int row,
            int col,
            out (int r1, int c1, int r2, int c2) bounds)
        {
            foreach (var mr in mergedRanges)
            {
                if (row >= mr.r1 && row <= mr.r2 && col >= mr.c1 && col <= mr.c2)
                {
                    bounds = mr;
                    return true;
                }
            }

            bounds = (row, col, row, col);
            return false;
        }

        private static string ToA1(int row1, int col1)
        {
            return ToColName(col1) + row1.ToString();
        }

        private static string ToA1Range(int r1, int c1, int r2, int c2)
        {
            var s = ToA1(r1, c1);
            var e = ToA1(r2, c2);
            return string.Equals(s, e, StringComparison.OrdinalIgnoreCase) ? s : s + ":" + e;
        }
        private static string ToColName(int col1)
        {
            if (col1 <= 0) return string.Empty;
            var chars = new Stack<char>();
            var n = col1;
            while (n > 0)
            {
                n--;
                chars.Push((char)('A' + (n % 26)));
                n /= 26;
            }
            return new string(chars.ToArray());
        }

        private static bool IsMeaningfulVisualCell(IXLCell cell, List<(int r1, int c1, int r2, int c2)> mergedRanges)
        {
            if (!HasMeaningfulValue(cell))
                return false;

            var row = cell.Address.RowNumber;
            var col = cell.Address.ColumnNumber;

            if (TryGetMergeBounds(mergedRanges, row, col, out var mr))
                return mr.r1 == row && mr.c1 == col;

            return true;
        }

        private static (int usedMaxRow, int usedMaxCol, string usedLastCellA1, List<(int r1, int c1, int r2, int c2)> merged) BuildVisibleBBox(IXLWorksheet ws)
        {
            var merged = ws.MergedRanges
                .Select(m => (
                    r1: m.RangeAddress.FirstAddress.RowNumber,
                    c1: m.RangeAddress.FirstAddress.ColumnNumber,
                    r2: m.RangeAddress.LastAddress.RowNumber,
                    c2: m.RangeAddress.LastAddress.ColumnNumber
                ))
                .ToList();

            var scanOptions = XLCellsUsedOptions.AllContents;

            int maxRow = 1;
            int maxCol = 1;
            bool found = false;

            foreach (var cell in ws.CellsUsed(scanOptions))
            {
                if (!HasMeaningfulValue(cell))
                    continue;

                var row = cell.Address.RowNumber;
                var col = cell.Address.ColumnNumber;

                if (TryGetMergeBounds(merged, row, col, out var mb))
                {
                    if (mb.r1 != row || mb.c1 != col)
                        continue;

                    if (mb.r2 > maxRow) maxRow = mb.r2;
                    if (mb.c2 > maxCol) maxCol = mb.c2;
                }
                else
                {
                    if (row > maxRow) maxRow = row;
                    if (col > maxCol) maxCol = col;
                }

                found = true;
            }

            if (!found)
                return (1, 1, "A1", merged);

            return (maxRow, maxCol, ToA1(maxRow, maxCol), merged);
        }


        private static string BuildStampRectsJsonFromExcel(
      string excelPath,
      string approvalCellsJson,
      string approvalsJson,
      string cooperationsJson,
      ILogger? log = null)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
                return "{\"widthPx\":0,\"heightPx\":0,\"stamps\":[]}";

            try
            {
                using var wb = new XLWorkbook(excelPath);
                var ws = wb.Worksheets.First();

                double defaultRowPt = ws.RowHeight <= 0 ? 15 : ws.RowHeight;
                double defaultColChar = ws.ColumnWidth <= 0 ? 8.43 : ws.ColumnWidth;

                int actualMaxC = 1;
                int actualMaxR = 1;

                foreach (var row in ws.RowsUsed(XLCellsUsedOptions.All))
                {
                    var rowNum = row.RowNumber();
                    if (rowNum > actualMaxR) actualMaxR = rowNum;

                    var last = row.LastCellUsed(XLCellsUsedOptions.All);
                    if (last != null && last.Address.ColumnNumber > actualMaxC)
                        actualMaxC = last.Address.ColumnNumber;
                }

                foreach (var mr in ws.MergedRanges)
                {
                    var lastR = mr.RangeAddress.LastAddress.RowNumber;
                    var lastC = mr.RangeAddress.LastAddress.ColumnNumber;
                    if (lastR > actualMaxR) actualMaxR = lastR;
                    if (lastC > actualMaxC) actualMaxC = lastC;
                }

                double ColPxAt(int col1)
                {
                    var w = ws.Column(col1).Width <= 0 ? defaultColChar : ws.Column(col1).Width;
                    return Math.Floor((w * 7d + 5d) / 7d) * 7d + 2d;
                }

                double RowPxAt(int row1)
                {
                    var h = ws.Row(row1).Height <= 0 ? defaultRowPt : ws.Row(row1).Height;
                    return Math.Round(h * 96d / 72d, 2);
                }

                double SumColsBefore(int col1)
                {
                    double sum = 0;
                    for (int c = 1; c < col1; c++) sum += ColPxAt(c);
                    return sum;
                }

                double SumRowsBefore(int row1)
                {
                    double sum = 0;
                    for (int r = 1; r < row1; r++) sum += RowPxAt(r);
                    return sum;
                }

                double sheetWidthPx = 0;
                for (int c = 1; c <= actualMaxC; c++) sheetWidthPx += ColPxAt(c);

                double sheetHeightPx = 0;
                for (int r = 1; r <= actualMaxR; r++) sheetHeightPx += RowPxAt(r);

                var merged = ws.MergedRanges
                    .Select(m => new
                    {
                        R1 = m.RangeAddress.FirstAddress.RowNumber,
                        C1 = m.RangeAddress.FirstAddress.ColumnNumber,
                        R2 = m.RangeAddress.LastAddress.RowNumber,
                        C2 = m.RangeAddress.LastAddress.ColumnNumber
                    })
                    .ToList();

                (int row, int col)? ParseA1(string? a1)
                {
                    var raw = (a1 ?? string.Empty).Trim().ToUpperInvariant();
                    var m = Regex.Match(raw, "^([A-Z]+)(\\d+)$");
                    if (!m.Success) return null;

                    int col = 0;
                    foreach (var ch in m.Groups[1].Value)
                        col = (col * 26) + (ch - 'A' + 1);

                    return (int.Parse(m.Groups[2].Value), col);
                }

                (int r1, int c1, int r2, int c2)? ParseA1Range(string? a1)
                {
                    var raw = (a1 ?? string.Empty).Trim().ToUpperInvariant();
                    if (string.IsNullOrWhiteSpace(raw)) return null;

                    if (!raw.Contains(':'))
                    {
                        var one = ParseA1(raw);
                        if (one == null) return null;
                        return (one.Value.row, one.Value.col, one.Value.row, one.Value.col);
                    }

                    var parts = raw.Split(':', StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length != 2) return null;

                    var s = ParseA1(parts[0]);
                    var e = ParseA1(parts[1]);
                    if (s == null || e == null) return null;

                    return (
                        Math.Min(s.Value.row, e.Value.row),
                        Math.Min(s.Value.col, e.Value.col),
                        Math.Max(s.Value.row, e.Value.row),
                        Math.Max(s.Value.col, e.Value.col)
                    );
                }

                (int r1, int c1, int r2, int c2) GetMergeBounds(int row1, int col1)
                {
                    foreach (var m in merged)
                    {
                        if (row1 >= m.R1 && row1 <= m.R2 && col1 >= m.C1 && col1 <= m.C2)
                            return (m.R1, m.C1, m.R2, m.C2);
                    }

                    return (row1, col1, row1, col1);
                }

                (int r1, int c1, int r2, int c2) UnionBounds(
                    (int r1, int c1, int r2, int c2) a,
                    (int r1, int c1, int r2, int c2) b)
                {
                    return (
                        Math.Min(a.r1, b.r1),
                        Math.Min(a.c1, b.c1),
                        Math.Max(a.r2, b.r2),
                        Math.Max(a.c2, b.c2)
                    );
                }

                string ToColName(int col1)
                {
                    if (col1 <= 0) return string.Empty;
                    var chars = new Stack<char>();
                    var n = col1;
                    while (n > 0)
                    {
                        n--;
                        chars.Push((char)('A' + (n % 26)));
                        n /= 26;
                    }
                    return new string(chars.ToArray());
                }

                string ToA1(int row1, int col1)
                {
                    return ToColName(col1) + row1.ToString();
                }

                string ToA1Range(int r1, int c1, int r2, int c2)
                {
                    var start = ToA1(r1, c1);
                    var end = ToA1(r2, c2);
                    return string.Equals(start, end, StringComparison.OrdinalIgnoreCase)
                        ? start
                        : start + ":" + end;
                }

                Dictionary<int, JsonElement> BuildApprovalMap(string rawJson)
                {
                    var map = new Dictionary<int, JsonElement>();
                    if (string.IsNullOrWhiteSpace(rawJson)) return map;

                    using var doc = JsonDocument.Parse(rawJson);
                    if (doc.RootElement.ValueKind != JsonValueKind.Array) return map;

                    foreach (var item in doc.RootElement.EnumerateArray())
                    {
                        int step = 0;

                        if (item.TryGetProperty("step", out var s1) && s1.ValueKind == JsonValueKind.Number) step = s1.GetInt32();
                        else if (item.TryGetProperty("Step", out var s2) && s2.ValueKind == JsonValueKind.Number) step = s2.GetInt32();

                        if (step > 0)
                            map[step] = item.Clone();
                    }

                    return map;
                }

                List<JsonElement> BuildArray(string rawJson)
                {
                    var list = new List<JsonElement>();
                    if (string.IsNullOrWhiteSpace(rawJson)) return list;

                    using var doc = JsonDocument.Parse(rawJson);
                    if (doc.RootElement.ValueKind != JsonValueKind.Array) return list;

                    foreach (var item in doc.RootElement.EnumerateArray())
                        list.Add(item.Clone());

                    return list;
                }

                string GetString(JsonElement el, params string[] names)
                {
                    foreach (var name in names)
                    {
                        if (el.TryGetProperty(name, out var p) && p.ValueKind == JsonValueKind.String)
                            return p.GetString() ?? string.Empty;
                    }
                    return string.Empty;
                }

                int GetInt(JsonElement el, params string[] names)
                {
                    foreach (var name in names)
                    {
                        if (el.TryGetProperty(name, out var p))
                        {
                            if (p.ValueKind == JsonValueKind.Number && p.TryGetInt32(out var n)) return n;
                            if (p.ValueKind == JsonValueKind.String && int.TryParse(p.GetString(), out var n2)) return n2;
                        }
                    }
                    return 0;
                }

                bool IsApproved(string actionRaw, string statusRaw)
                {
                    return actionRaw == "APPROVE"
                        || actionRaw.StartsWith("APPROV", StringComparison.OrdinalIgnoreCase)
                        || statusRaw == "APPROVED"
                        || statusRaw.StartsWith("APPROVED", StringComparison.OrdinalIgnoreCase)
                        || statusRaw == "COOPERATED"
                        || statusRaw.StartsWith("COOPERATED", StringComparison.OrdinalIgnoreCase);
                }

                bool IsHold(string actionRaw, string statusRaw)
                {
                    return actionRaw == "HOLD"
                        || actionRaw.StartsWith("HOLD", StringComparison.OrdinalIgnoreCase)
                        || statusRaw == "HOLD"
                        || statusRaw == "ONHOLD"
                        || statusRaw.StartsWith("PENDINGHOLD", StringComparison.OrdinalIgnoreCase)
                        || statusRaw.StartsWith("HOLD", StringComparison.OrdinalIgnoreCase);
                }

                bool IsRejected(string actionRaw, string statusRaw)
                {
                    return actionRaw == "REJECT"
                        || actionRaw.StartsWith("REJECT", StringComparison.OrdinalIgnoreCase)
                        || statusRaw == "REJECTED"
                        || statusRaw == "REJECT"
                        || statusRaw.StartsWith("REJECT", StringComparison.OrdinalIgnoreCase);
                }

                string GetStampKind(string actionRaw, string statusRaw)
                {
                    if (IsApproved(actionRaw, statusRaw)) return "approved";
                    if (IsHold(actionRaw, statusRaw)) return "holding";
                    if (IsRejected(actionRaw, statusRaw)) return "rejected";
                    return string.Empty;
                }

                var approvalMapRows = BuildArray(approvalCellsJson);
                var approvalsByStep = BuildApprovalMap(approvalsJson);
                var cooperationRows = BuildArray(cooperationsJson);

                var stamps = new List<object>();

                foreach (var cellRow in approvalMapRows)
                {
                    var step = GetInt(cellRow, "Step", "step", "Slot", "slot");
                    var a1 = GetString(cellRow, "A1", "a1", "CellA1", "cellA1");
                    if (step <= 0 || string.IsNullOrWhiteSpace(a1)) continue;
                    if (!approvalsByStep.TryGetValue(step, out var approval)) continue;

                    var actionRaw = GetString(approval, "action", "Action").Trim().ToUpperInvariant();
                    var statusRaw = GetString(approval, "status", "Status").Trim().ToUpperInvariant();
                    var stampKind = GetStampKind(actionRaw, statusRaw);
                    if (string.IsNullOrWhiteSpace(stampKind)) continue;

                    var parsed = ParseA1Range(a1);
                    if (parsed == null) continue;

                    var b1 = GetMergeBounds(parsed.Value.r1, parsed.Value.c1);
                    var b2 = GetMergeBounds(parsed.Value.r2, parsed.Value.c2);
                    var bounds = UnionBounds(b1, b2);
                    var boundsA1 = ToA1Range(bounds.r1, bounds.c1, bounds.r2, bounds.c2);

                    double x = SumColsBefore(bounds.c1);
                    double y = SumRowsBefore(bounds.r1);
                    double w = 0;
                    double h = 0;

                    for (int c = bounds.c1; c <= bounds.c2; c++) w += ColPxAt(c);
                    for (int r = bounds.r1; r <= bounds.r2; r++) h += RowPxAt(r);

                    stamps.Add(new
                    {
                        lineType = "approval",
                        step,
                        roleKey = GetString(approval, "roleKey", "RoleKey"),
                        a1,
                        boundsA1,
                        row1 = bounds.r1,
                        col1 = bounds.c1,
                        row2 = bounds.r2,
                        col2 = bounds.c2,
                        x = Math.Round(x, 2),
                        y = Math.Round(y, 2),
                        w = Math.Round(w, 2),
                        h = Math.Round(h, 2),
                        action = GetString(approval, "action", "Action"),
                        status = GetString(approval, "status", "Status"),
                        approverDisplayText = GetString(approval, "approverDisplayText", "ApproverDisplayText"),
                        signaturePath = GetString(approval, "signaturePath", "SignaturePath"),
                        stampKind
                    });
                }

                foreach (var coop in cooperationRows)
                {
                    var a1 = GetString(coop, "cellA1", "CellA1", "A1", "a1");
                    if (string.IsNullOrWhiteSpace(a1)) continue;

                    var actionRaw = GetString(coop, "action", "Action").Trim().ToUpperInvariant();
                    var statusRaw = GetString(coop, "status", "Status").Trim().ToUpperInvariant();
                    var stampKind = GetStampKind(actionRaw, statusRaw);
                    if (string.IsNullOrWhiteSpace(stampKind)) continue;

                    var parsed = ParseA1Range(a1);
                    if (parsed == null) continue;

                    var b1 = GetMergeBounds(parsed.Value.r1, parsed.Value.c1);
                    var b2 = GetMergeBounds(parsed.Value.r2, parsed.Value.c2);
                    var bounds = UnionBounds(b1, b2);
                    var boundsA1 = ToA1Range(bounds.r1, bounds.c1, bounds.r2, bounds.c2);

                    double x = SumColsBefore(bounds.c1);
                    double y = SumRowsBefore(bounds.r1);
                    double w = 0;
                    double h = 0;

                    for (int c = bounds.c1; c <= bounds.c2; c++) w += ColPxAt(c);
                    for (int r = bounds.r1; r <= bounds.r2; r++) h += RowPxAt(r);

                    stamps.Add(new
                    {
                        lineType = "cooperation",
                        roleKey = GetString(coop, "roleKey", "RoleKey"),
                        a1,
                        boundsA1,
                        row1 = bounds.r1,
                        col1 = bounds.c1,
                        row2 = bounds.r2,
                        col2 = bounds.c2,
                        x = Math.Round(x, 2),
                        y = Math.Round(y, 2),
                        w = Math.Round(w, 2),
                        h = Math.Round(h, 2),
                        action = GetString(coop, "action", "Action"),
                        status = GetString(coop, "status", "Status"),
                        approverDisplayText = GetString(coop, "approverDisplayText", "ApproverDisplayText"),
                        signaturePath = GetString(coop, "signaturePath", "SignaturePath"),
                        stampKind
                    });
                }

                return JsonSerializer.Serialize(new
                {
                    widthPx = Math.Round(sheetWidthPx, 2),
                    heightPx = Math.Round(sheetHeightPx, 2),
                    stamps
                });
            }
            catch (Exception ex)
            {
                log?.LogWarning(ex, "BuildStampRectsJsonFromExcel failed path={Path}", excelPath);
                return "{\"widthPx\":0,\"heightPx\":0,\"stamps\":[]}";
            }
        }

        // ========== ApproveOrHold ==========
        [HttpPost("ApproveOrHoldDX")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> ApproveOrHold([FromBody] ApproveDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto?.docId) || string.IsNullOrWhiteSpace(dto?.action))
                return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" } });

            var actionLower = dto.action!.ToLowerInvariant();

            static bool IsHoldLikeStatus(string? st)
            {
                if (string.IsNullOrWhiteSpace(st)) return false;
                st = st.Trim();
                if (st.StartsWith("PendingHold", StringComparison.OrdinalIgnoreCase)) return true;
                if (st.Equals("OnHold", StringComparison.OrdinalIgnoreCase)) return true;
                if (st.Equals("Hold", StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            }

            // ── 회수 ──────────────────────────────────────────────────────────
            if (actionLower == "recall")
            {
                var csR = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(csR);
                await conn.OpenAsync();

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
                    ac.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = _helper.GetCurrentUserDisplayNameStrict() ?? string.Empty });
                    await ac.ExecuteNonQueryAsync();
                }

                return Json(new { ok = true, docId = dto.docId, status = "Recalled" });
            }

            // ── 협조 확인
            if (actionLower == "cooperate")
            {
                var approverId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
                var approverName = _helper.GetCurrentUserDisplayNameStrict();

                string? coopDisplayText = null;
                string? coopSignaturePath = null;
                try
                {
                    var csP = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                    await using var connP = new SqlConnection(csP);
                    await connP.OpenAsync();
                    await using var up = connP.CreateCommand();
                    up.CommandText = @"
SELECT TOP (1)
       u.DisplayName,
       COALESCE(pl.Name, pm.Name, N'') AS PositionName,
       COALESCE(dl.Name, dm.Name, N'') AS DepartmentName,
       u.SignatureRelativePath         AS SignaturePath
FROM dbo.UserProfiles AS u
LEFT JOIN dbo.DepartmentMasters   AS dm ON dm.CompCd = u.CompCd AND dm.Id       = u.DepartmentId
LEFT JOIN dbo.DepartmentMasterLoc AS dl ON dl.DepartmentId = dm.Id AND dl.LangCode = N'ko'
LEFT JOIN dbo.PositionMasters     AS pm ON pm.CompCd = u.CompCd AND pm.Id       = u.PositionId
LEFT JOIN dbo.PositionMasterLoc   AS pl ON pl.PositionId   = pm.Id AND pl.LangCode = N'ko'
WHERE u.UserId = @uid;";
                    up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });
                    await using var r = await up.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        var disp = r["DisplayName"] as string ?? string.Empty;
                        var pos = r["PositionName"] as string ?? string.Empty;
                        var dept = r["DepartmentName"] as string ?? string.Empty;
                        var sig = r["SignaturePath"] as string ?? string.Empty;
                        coopDisplayText = string.Join(" ", new[] { dept, pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));
                        coopSignaturePath = string.IsNullOrWhiteSpace(sig) ? null : sig;
                    }
                }
                catch (Exception ex)
                {
                    _log.LogError(ex, "cooperate: profile snapshot failed docId={docId}", dto.docId);
                }

                var csC = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(csC);
                await conn.OpenAsync();

                int? coopId = null;
                await using (var ck = conn.CreateCommand())
                {
                    ck.CommandText = @"
SELECT TOP 1 Id
FROM dbo.DocumentCooperations
WHERE DocId  = @DocId
  AND UserId = @UserId
  AND ISNULL(Status, N'Pending') = N'Pending';";
                    ck.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    ck.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId });
                    var o = await ck.ExecuteScalarAsync();
                    if (o == null || o == DBNull.Value)
                        return Forbid();
                    coopId = Convert.ToInt32(o);
                }

                await using (var upC = conn.CreateCommand())
                {
                    upC.CommandText = @"
UPDATE dbo.DocumentCooperations
SET Status              = N'Cooperated',
    Action              = N'Cooperate',
    ActedAt             = SYSUTCDATETIME(),
    ActorName           = @actor,
    ApproverDisplayText = COALESCE(@displayText, ApproverDisplayText)
WHERE Id = @Id;";
                    upC.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int) { Value = coopId!.Value });
                    upC.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName ?? string.Empty });
                    upC.Parameters.Add(new SqlParameter("@displayText", SqlDbType.NVarChar, 400) { Value = (object?)coopDisplayText ?? DBNull.Value });
                    await upC.ExecuteNonQueryAsync();
                }
                ///
                try
                {
                    var csChk = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                    await using var connChk = new SqlConnection(csChk);
                    await connChk.OpenAsync();

                    // 현재 문서 상태가 PendingA{N} 형태인지 확인
                    string? docStatusNow = null;
                    await using (var dsChk = connChk.CreateCommand())
                    {
                        dsChk.CommandText = @"SELECT TOP 1 Status FROM dbo.Documents WHERE DocId = @DocId;";
                        dsChk.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        docStatusNow = await dsChk.ExecuteScalarAsync() as string;
                    }

                    // PendingA{N} 패턴 파싱
                    int? pendingStep = null;
                    if (!string.IsNullOrWhiteSpace(docStatusNow)
                        && docStatusNow.StartsWith("PendingA", StringComparison.OrdinalIgnoreCase))
                    {
                        if (int.TryParse(docStatusNow.Substring("PendingA".Length), out var ps))
                            pendingStep = ps;
                    }

                    if (pendingStep != null)
                    {
                        // pendingStep이 마지막 결재자인지 확인
                        int? stepAfterPending = null;
                        await using (var chkL = connChk.CreateCommand())
                        {
                            chkL.CommandText = @"
SELECT TOP 1 StepOrder
FROM dbo.DocumentApprovals
WHERE DocId = @DocId AND StepOrder > @Step
ORDER BY StepOrder;";
                            chkL.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            chkL.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = pendingStep.Value });
                            var o2 = await chkL.ExecuteScalarAsync();
                            if (o2 != null && o2 != DBNull.Value) stepAfterPending = Convert.ToInt32(o2);
                        }

                        bool isPendingFinal = (stepAfterPending == null);

                        if (isPendingFinal)
                        {
                            // 남은 협조자 수 확인
                            await using var chkCoop2 = connChk.CreateCommand();
                            chkCoop2.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentCooperations
WHERE DocId = @DocId
  AND ISNULL(Status, N'Pending') NOT IN (N'Cooperated', N'Recalled');";
                            chkCoop2.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            var remaining = Convert.ToInt32(await chkCoop2.ExecuteScalarAsync());

                            if (remaining == 0)
                            {
                                // 모든 협조 완료 → 마지막 결재자에게 알림
                                var finalIds = await GetNextApproverUserIdsFromDbAsync(connChk, dto.docId!, pendingStep.Value);
                                await DocControllerHelper.SendApprovalPendingBadgeAsync(
                                    notifier: _webPushNotifier, S: _S, conn: connChk,
                                    targetUserIds: finalIds ?? new List<string>(),
                                    url: "/", tag: "badge-approval-pending");
                            }
                        }
                    }
                }
                catch (Exception exCoopPush)
                {
                    _log.LogWarning(exCoopPush, "cooperate: final step notify failed docId={docId}", dto.docId);
                }

                ///

                return Json(new { ok = true, docId = dto.docId, status = "Cooperated" });
            }


            // ── 협조 거부
            if (actionLower == "cooperatereject")
            {
                var approverId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
                var approverName = _helper.GetCurrentUserDisplayNameStrict();

                string? coopDisplayText = null;
                try
                {
                    var csP = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                    await using var connP = new SqlConnection(csP);
                    await connP.OpenAsync();
                    await using var up = connP.CreateCommand();
                    up.CommandText = @"
SELECT TOP (1)
       u.DisplayName,
       COALESCE(pl.Name, pm.Name, N'') AS PositionName,
       COALESCE(dl.Name, dm.Name, N'') AS DepartmentName
FROM dbo.UserProfiles AS u
LEFT JOIN dbo.DepartmentMasters   AS dm ON dm.CompCd = u.CompCd AND dm.Id       = u.DepartmentId
LEFT JOIN dbo.DepartmentMasterLoc AS dl ON dl.DepartmentId = dm.Id AND dl.LangCode = N'ko'
LEFT JOIN dbo.PositionMasters     AS pm ON pm.CompCd = u.CompCd AND pm.Id       = u.PositionId
LEFT JOIN dbo.PositionMasterLoc   AS pl ON pl.PositionId   = pm.Id AND pl.LangCode = N'ko'
WHERE u.UserId = @uid;";
                    up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });

                    await using var r = await up.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        var disp = r["DisplayName"] as string ?? string.Empty;
                        var pos = r["PositionName"] as string ?? string.Empty;
                        var dept = r["DepartmentName"] as string ?? string.Empty;
                        coopDisplayText = string.Join(" ", new[] { dept, pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));
                    }
                }
                catch (Exception ex)
                {
                    _log.LogError(ex, "cooperateReject: profile snapshot failed docId={docId}", dto.docId);
                }

                var csC = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(csC);
                await conn.OpenAsync();

                int? coopId = null;
                await using (var ck = conn.CreateCommand())
                {
                    ck.CommandText = @"
SELECT TOP 1 Id
FROM dbo.DocumentCooperations
WHERE DocId  = @DocId
  AND UserId = @UserId
  AND ISNULL(Status, N'Pending') = N'Pending';";
                    ck.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    ck.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId });

                    var o = await ck.ExecuteScalarAsync();
                    if (o == null || o == DBNull.Value)
                        return Forbid();

                    coopId = Convert.ToInt32(o);
                }

                await using (var upC = conn.CreateCommand())
                {
                    upC.CommandText = @"
UPDATE dbo.DocumentCooperations
SET Status              = N'Rejected',
    Action              = N'Reject',
    ActedAt             = SYSUTCDATETIME(),
    ActorName           = @actor,
    ApproverDisplayText = COALESCE(@displayText, ApproverDisplayText)
WHERE Id = @Id;";
                    upC.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int) { Value = coopId!.Value });
                    upC.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName ?? string.Empty });
                    upC.Parameters.Add(new SqlParameter("@displayText", SqlDbType.NVarChar, 300) { Value = (object?)coopDisplayText ?? DBNull.Value });

                    await upC.ExecuteNonQueryAsync();
                }

                return Json(new
                {
                    ok = true,
                    docId = dto.docId,
                    status = "Rejected"
                });
            }
            // ── 승인 / 보류 / 반려 
            string outXlsx = string.Empty;
            try
            {
                var csOut = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var connOut = new SqlConnection(csOut);
                await connOut.OpenAsync();
                await using var cmdOut = connOut.CreateCommand();
                cmdOut.CommandText = @"SELECT TOP (1) OutputPath FROM dbo.Documents WHERE DocId = @DocId;";
                cmdOut.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                var obj = await cmdOut.ExecuteScalarAsync();
                outXlsx = (obj as string) ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(outXlsx) && !System.IO.Path.IsPathRooted(outXlsx))
                    outXlsx = System.IO.Path.Combine(_env.ContentRootPath, outXlsx.TrimStart('\\', '/'));
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold: output path resolve failed docId={docId}", dto.docId);
            }

            if (string.IsNullOrWhiteSpace(outXlsx) || !System.IO.File.Exists(outXlsx))
                return NotFound(new { messages = new[] { "DOC_Err_DocumentNotFound" } });

            var approverId2 = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            var approverName2 = _helper.GetCurrentUserDisplayNameStrict();

            string? approverDisplayText = null;
            string? signatureRelativePath = null;

            try
            {
                var csP = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var connP = new SqlConnection(csP);
                await connP.OpenAsync();
                await using var up = connP.CreateCommand();
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
                up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId2 });

                await using var r = await up.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    var disp = r["DisplayName"] as string ?? string.Empty;
                    var pos = r["PositionName"] as string ?? string.Empty;
                    var dept = r["DepartmentName"] as string ?? string.Empty;
                    var sig = r["SignaturePath"] as string ?? string.Empty;
                    approverDisplayText = string.Join(" ", new[] { dept, pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));
                    signatureRelativePath = string.IsNullOrWhiteSpace(sig) ? null : sig;
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold: profile snapshot failed docId={docId}", dto.docId);
            }

            string newStatus = "Updated";
            int? actedStep = null;
            int? nextStepForNotify = null;

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                string docStatus = string.Empty;
                await using (var ds = conn.CreateCommand())
                {
                    ds.CommandText = @"SELECT TOP (1) Status FROM dbo.Documents WHERE DocId = @DocId;";
                    ds.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    var o = await ds.ExecuteScalarAsync();
                    docStatus = (o as string) ?? string.Empty;
                }

                int? currentStep = null;
                await using (var findStep = conn.CreateCommand())
                {
                    findStep.CommandText = @"
SELECT TOP (1) StepOrder
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND ISNULL(Status,N'Pending') LIKE N'Pending%'
  AND (UserId = @UserId OR ApproverValue = @UserId)
ORDER BY StepOrder;";
                    findStep.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    findStep.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId2 });
                    var stepObj = await findStep.ExecuteScalarAsync();
                    if (stepObj != null && stepObj != DBNull.Value)
                        currentStep = Convert.ToInt32(stepObj);
                }

                if (currentStep == null)
                    return Forbid();

                var step = currentStep.Value;
                actedStep = step;

                if (actionLower == "hold")
                {
                    string apprStatus = string.Empty;
                    await using (var asel = conn.CreateCommand())
                    {
                        asel.CommandText = @"
SELECT TOP (1) Status
FROM dbo.DocumentApprovals
WHERE DocId = @DocId AND StepOrder = @Step;";
                        asel.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        asel.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = step });
                        var ao = await asel.ExecuteScalarAsync();
                        apprStatus = (ao as string) ?? string.Empty;
                    }
                    if (IsHoldLikeStatus(apprStatus) || IsHoldLikeStatus(docStatus))
                        return Conflict(new { messages = new[] { "DOC_Err_BadRequest" } });
                }

                await using (var u = conn.CreateCommand())
                {
                    u.CommandText = @"
UPDATE dbo.DocumentApprovals
SET Action              = @act,
    ActedAt             = SYSUTCDATETIME(),
    ActorName           = @actor,
    Status              = CASE
                             WHEN @act = N'approve' THEN N'Approved'
                             WHEN @act = N'hold'    THEN N'PendingHold'
                             WHEN @act = N'reject'  THEN N'Rejected'
                             ELSE ISNULL(Status, N'Updated')
                          END,
    UserId              = COALESCE(UserId, @uid),
    ApproverDisplayText = COALESCE(@displayText, ApproverDisplayText),
    SignaturePath       = COALESCE(@sigPath, SignaturePath)
WHERE DocId = @id AND StepOrder = @step;";
                    u.Parameters.Add(new SqlParameter("@act", SqlDbType.NVarChar, 20) { Value = actionLower });
                    u.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName2 ?? string.Empty });
                    u.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId2 });
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    u.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = step });
                    u.Parameters.Add(new SqlParameter("@displayText", SqlDbType.NVarChar, 400) { Value = (object?)approverDisplayText ?? DBNull.Value });
                    u.Parameters.Add(new SqlParameter("@sigPath", SqlDbType.NVarChar, 400) { Value = (object?)signatureRelativePath ?? DBNull.Value });
                    await u.ExecuteNonQueryAsync();
                }

                newStatus = actionLower switch
                {
                    "approve" => $"ApprovedA{step}",
                    "hold" => $"PendingHoldA{step}",
                    "reject" => $"RejectedA{step}",
                    _ => "Updated"
                };

                await using (var u2 = conn.CreateCommand())
                {
                    u2.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                    u2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = newStatus });
                    u2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    await u2.ExecuteNonQueryAsync();
                }

                if (actionLower == "approve")
                {
                    var next = step + 1;
                    nextStepForNotify = next;
                    var toList = await GetNextApproverEmailsFromDbAsync(conn, dto.docId!, next);

                    if (toList.Count == 0)
                    {
                        // 다음 결재자 없음 → 최종 승인
                        await using var up2a = conn.CreateCommand();
                        up2a.CommandText = @"UPDATE dbo.Documents SET Status = N'Approved' WHERE DocId = @id;";
                        up2a.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        await up2a.ExecuteNonQueryAsync();
                        newStatus = "Approved";
                    }
                    else
                    {
                        // ── next가 마지막 결재자인지 확인 ──
                        int? stepAfterNext = null;
                        await using (var chkLast = conn.CreateCommand())
                        {
                            chkLast.CommandText = @"
SELECT TOP 1 StepOrder
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder > @Next
ORDER BY StepOrder;";
                            chkLast.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            chkLast.Parameters.Add(new SqlParameter("@Next", SqlDbType.Int) { Value = next });
                            var o = await chkLast.ExecuteScalarAsync();
                            if (o != null && o != DBNull.Value) stepAfterNext = Convert.ToInt32(o);
                        }

                        bool isNextFinalStep = (stepAfterNext == null); // next가 마지막 결재자

                        bool coopAllDone = true;
                        if (isNextFinalStep)
                        {
                            // 협조자 중 Cooperated 아닌 사람 있는지 확인
                            await using var chkCoop = conn.CreateCommand();
                            chkCoop.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentCooperations
WHERE DocId = @DocId
  AND ISNULL(Status, N'Pending') NOT IN (N'Cooperated', N'Recalled');";
                            chkCoop.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            var coopPending = Convert.ToInt32(await chkCoop.ExecuteScalarAsync());
                            coopAllDone = (coopPending == 0);
                        }

                        if (isNextFinalStep && !coopAllDone)
                        {
                            // 협조 미완료 → 마지막 결재자에게 아직 알림 보내지 않고 대기 상태만 기록
                            await using (var up2 = conn.CreateCommand())
                            {
                                up2.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                                up2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = $"PendingA{next}" });
                                up2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                                await up2.ExecuteNonQueryAsync();
                            }
                            newStatus = $"PendingA{next}";
                            // 알림 발송 안 함 — 협조 완료 시 cooperate 쪽에서 발송
                        }
                        else
                        {
                            // 협조자 없거나 협조 완료 → 바로 알림 발송
                            await using (var up2 = conn.CreateCommand())
                            {
                                up2.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                                up2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = $"PendingA{next}" });
                                up2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                                await up2.ExecuteNonQueryAsync();
                            }
                            newStatus = $"PendingA{next}";

                            try
                            {
                                var nextIds = await GetNextApproverUserIdsFromDbAsync(conn, dto.docId!, next);
                                await DocControllerHelper.SendApprovalPendingBadgeAsync(
                                    notifier: _webPushNotifier, S: _S, conn: conn,
                                    targetUserIds: nextIds ?? new List<string>(),
                                    url: "/", tag: "badge-approval-pending");
                            }
                            catch (Exception exPush)
                            {
                                _log.LogWarning(exPush, "ApproveOrHold: push notify failed docId={docId}", dto.docId);
                            }
                        }
                    }
                }
                else if (actionLower == "hold" || actionLower == "reject")
                {
                    try
                    {
                        var authorId = await GetDocumentAuthorUserIdAsync(conn, dto.docId!);
                        if (!string.IsNullOrWhiteSpace(authorId)
                            && !string.Equals(authorId, approverId2, StringComparison.OrdinalIgnoreCase))
                        {
                            var titleText = (_S?["PUSH_SummaryTitle"] ?? "PUSH_SummaryTitle").ToString();
                            var bodyText = actionLower == "hold"
                                ? (_S?["PUSH_ApprovalHold"] ?? "PUSH_ApprovalHold").ToString()
                                : (_S?["PUSH_ApprovalReject"] ?? "PUSH_ApprovalReject").ToString();
                            var tag = actionLower == "hold" ? "approval-author-hold" : "approval-author-reject";
                            await _webPushNotifier.SendToUserIdAsync(
                                userId: authorId.Trim(), title: titleText, body: bodyText,
                                url: "/Doc/DetailDX?id=" + Uri.EscapeDataString(dto.docId!), tag: tag);
                        }
                    }
                    catch (Exception exPush)
                    {
                        _log.LogWarning(exPush, "ApproveOrHold: push notify(author) failed docId={docId}", dto.docId);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold flow failed docId={docId}", dto.docId);
            }

            var previewJson2 = BuildPreviewJsonFromExcel(outXlsx);
            return Json(new { ok = true, docId = dto.docId, status = newStatus, previewJson = previewJson2 });
        }

//        // ========== UpdateShares ==========
//        [HttpPost("UpdateShares")]
//        [ValidateAntiForgeryToken]
//        [Consumes("application/json")]
//        [Produces("application/json")]
//        public async Task<IActionResult> UpdateShares([FromBody] UpdateSharesDto dto)
//        {
//            if (dto is null || string.IsNullOrWhiteSpace(dto.DocId))
//                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_SaveFailed" }, stage = "arg" });

//            var docId = dto.DocId!.Trim();
//            var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

//            var selected = (dto.SelectedRecipientUserIds ?? new List<string>())
//                .Where(x => !string.IsNullOrWhiteSpace(x))
//                .Select(x => x.Trim())
//                .Distinct(StringComparer.OrdinalIgnoreCase)
//                .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
//                .ToList();

//            var newlyAdded = new List<string>();

//            try
//            {
//                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
//                await using var conn = new SqlConnection(cs);
//                await conn.OpenAsync();
//                await using var tx = await conn.BeginTransactionAsync();

//                try
//                {
//                    var currentActive = new List<string>();
//                    await using (var sel = conn.CreateCommand())
//                    {
//                        sel.Transaction = (SqlTransaction)tx;
//                        sel.CommandText = @"SELECT UserId FROM dbo.DocumentShares WHERE DocId = @DocId AND ISNULL(IsRevoked,0) = 0;";
//                        sel.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
//                        await using var r = await sel.ExecuteReaderAsync();
//                        while (await r.ReadAsync())
//                        {
//                            var uid = (r["UserId"]?.ToString() ?? string.Empty).Trim();
//                            if (!string.IsNullOrWhiteSpace(uid)) currentActive.Add(uid);
//                        }
//                    }

//                    newlyAdded = selected
//                        .Where(uid => !currentActive.Any(x => string.Equals(x, uid, StringComparison.OrdinalIgnoreCase)))
//                        .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

//                    var toRevoke = currentActive
//                        .Where(uid => !selected.Any(x => string.Equals(x, uid, StringComparison.OrdinalIgnoreCase)))
//                        .ToList();

//                    foreach (var targetUserId in toRevoke)
//                    {
//                        await using (var cmd = conn.CreateCommand())
//                        {
//                            cmd.Transaction = (SqlTransaction)tx;
//                            cmd.CommandText = @"UPDATE dbo.DocumentShares SET IsRevoked = 1, ExpireAt = SYSUTCDATETIME() WHERE DocId = @DocId AND UserId = @UserId;";
//                            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
//                            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
//                            await cmd.ExecuteNonQueryAsync();
//                        }
//                        await using (var logCmd = conn.CreateCommand())
//                        {
//                            logCmd.Transaction = (SqlTransaction)tx;
//                            logCmd.CommandText = @"INSERT INTO dbo.DocumentShareLogs (DocId,ActorId,ChangeCode,TargetUserId,BeforeJson,AfterJson,ChangedAt) VALUES (@DocId,@ActorId,@ChangeCode,@TargetUserId,NULL,@AfterJson,SYSUTCDATETIME());";
//                            logCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
//                            logCmd.Parameters.Add(new SqlParameter("@ActorId", SqlDbType.NVarChar, 64) { Value = actorId });
//                            logCmd.Parameters.Add(new SqlParameter("@ChangeCode", SqlDbType.NVarChar, 50) { Value = "ShareRevoked" });
//                            logCmd.Parameters.Add(new SqlParameter("@TargetUserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
//                            logCmd.Parameters.Add(new SqlParameter("@AfterJson", SqlDbType.NVarChar, -1) { Value = "{\"revoked\":true}" });
//                            await logCmd.ExecuteNonQueryAsync();
//                        }
//                    }

//                    foreach (var targetUserId in selected)
//                    {
//                        await using (var cmd = conn.CreateCommand())
//                        {
//                            cmd.Transaction = (SqlTransaction)tx;
//                            cmd.CommandText = @"
//IF EXISTS (SELECT 1 FROM dbo.DocumentShares WHERE DocId=@DocId AND UserId=@UserId)
//BEGIN
//    UPDATE dbo.DocumentShares SET IsRevoked = 0, ExpireAt = NULL WHERE DocId=@DocId AND UserId=@UserId;
//END
//ELSE
//BEGIN
//    INSERT INTO dbo.DocumentShares (DocId,UserId,AccessRole,ExpireAt,IsRevoked,CreatedBy,CreatedAt)
//    VALUES (@DocId,@UserId,'Commenter',NULL,0,@CreatedBy,SYSUTCDATETIME());
//END";
//                            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
//                            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
//                            cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 64) { Value = actorId });
//                            await cmd.ExecuteNonQueryAsync();
//                        }
//                        await using (var logCmd = conn.CreateCommand())
//                        {
//                            logCmd.Transaction = (SqlTransaction)tx;
//                            logCmd.CommandText = @"INSERT INTO dbo.DocumentShareLogs (DocId,ActorId,ChangeCode,TargetUserId,BeforeJson,AfterJson,ChangedAt) VALUES (@DocId,@ActorId,@ChangeCode,@TargetUserId,NULL,@AfterJson,SYSUTCDATETIME());";
//                            logCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
//                            logCmd.Parameters.Add(new SqlParameter("@ActorId", SqlDbType.NVarChar, 64) { Value = actorId });
//                            logCmd.Parameters.Add(new SqlParameter("@ChangeCode", SqlDbType.NVarChar, 50) { Value = "ShareAdded" });
//                            logCmd.Parameters.Add(new SqlParameter("@TargetUserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
//                            logCmd.Parameters.Add(new SqlParameter("@AfterJson", SqlDbType.NVarChar, -1) { Value = "{\"accessRole\":\"Commenter\"}" });
//                            await logCmd.ExecuteNonQueryAsync();
//                        }
//                    }

//                    await ((SqlTransaction)tx).CommitAsync();
//                }
//                catch
//                {
//                    await ((SqlTransaction)tx).RollbackAsync();
//                    throw;
//                }

//                try
//                {
//                    var ids = newlyAdded.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
//                    if (ids.Count > 0)
//                    {
//                        await using var connPush = new SqlConnection(_cfg.GetConnectionString("DefaultConnection") ?? string.Empty);
//                        await connPush.OpenAsync();
//                        await DocControllerHelper.SendSharedUnreadBadgeAsync(
//                            notifier: _webPushNotifier, S: _S, conn: connPush,
//                            targetUserIds: ids, url: "/", tag: "badge-shared");
//                    }
//                }
//                catch (Exception exPush)
//                {
//                    _log.LogWarning(exPush, "UpdateShares: push notify failed docId={docId}", docId);
//                }

//                return Json(new { ok = true, docId, selectedCount = selected.Count, addedNotifyCount = newlyAdded.Count });
//            }
//            catch (Exception ex)
//            {
//                _log.LogWarning(ex, "UpdateShares failed docId={docId}", docId);
//                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_SaveFailed" }, stage = "db", detail = ex.Message });
//            }
//        }

        // ========== DetailsData ==========
        [HttpGet("DetailsData")]
        [Produces("application/json")]
        public async Task<IActionResult> DetailsData(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return BadRequest(new { message = "DOC_Err_BadRequest" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT d.DocId, d.TemplateCode, d.TemplateTitle, d.Status, d.DescriptorJson, d.OutputPath, d.CreatedAt
FROM dbo.Documents d WHERE d.DocId = @id;";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

            string? descriptor = null, output = null, title = null, code = null, status = null;
            DateTime? createdAt = null;

            await using (var rd = await cmd.ExecuteReaderAsync())
            {
                if (await rd.ReadAsync())
                {
                    code = rd["TemplateCode"] as string ?? string.Empty;
                    title = rd["TemplateTitle"] as string ?? string.Empty;
                    status = rd["Status"] as string ?? string.Empty;
                    descriptor = rd["DescriptorJson"] as string ?? "{}";
                    output = rd["OutputPath"] as string ?? string.Empty;
                    if (!rd.IsDBNull(rd.GetOrdinal("CreatedAt")))
                        createdAt = (DateTime)rd["CreatedAt"];
                }
                else return NotFound(new { message = "DOC_Err_DocumentNotFound" });
            }

            var resolvedOutput = output ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(resolvedOutput) && !System.IO.Path.IsPathRooted(resolvedOutput))
                resolvedOutput = System.IO.Path.Combine(_env.ContentRootPath, resolvedOutput.TrimStart('\\', '/'));

            var preview = string.IsNullOrWhiteSpace(resolvedOutput) || !System.IO.File.Exists(resolvedOutput)
                ? "{}" : BuildPreviewJsonFromExcel(resolvedOutput);

            var approvals = new List<object>();
            await using (var ac = conn.CreateCommand())
            {
                ac.CommandText = @"
SELECT a.StepOrder, a.RoleKey, a.ApproverValue, a.UserId, a.Status, a.Action, a.ActedAt, a.ActorName,
       COALESCE(a.ApproverDisplayText, up.DisplayName, a.ActorName, a.ApproverValue) AS ApproverDisplayText
FROM dbo.DocumentApprovals a
LEFT JOIN dbo.UserProfiles up ON up.UserId = a.UserId
WHERE a.DocId = @DocId ORDER BY a.StepOrder;";
                ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });
                await using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    DateTime? acted = r.IsDBNull(6) ? (DateTime?)null : r.GetDateTime(6);
                    approvals.Add(new
                    {
                        step = r.GetInt32(0),
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue = r["ApproverValue"]?.ToString(),
                        userId = r["UserId"]?.ToString(),
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = acted.HasValue ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc)) : null,
                        actorName = r["ActorName"]?.ToString(),
                        ApproverDisplayText = r["ApproverDisplayText"]?.ToString() ?? string.Empty
                    });
                }
            }

            var cooperations = new List<object>();
            await using (var cc = conn.CreateCommand())
            {
                cc.CommandText = @"
SELECT dc.RoleKey, dc.UserId, dc.ApproverValue, dc.Status, dc.Action, dc.ActedAt, dc.ActorName,
       COALESCE(dc.ApproverDisplayText, up.DisplayName, dc.ActorName, dc.ApproverValue) AS ApproverDisplayText
FROM dbo.DocumentCooperations dc
LEFT JOIN dbo.UserProfiles up ON up.UserId = dc.UserId
WHERE dc.DocId = @DocId
ORDER BY dc.RoleKey;";
                cc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });
                await using var cr = await cc.ExecuteReaderAsync();
                while (await cr.ReadAsync())
                {
                    DateTime? acted = cr["ActedAt"] is DateTime dt ? dt : (DateTime?)null;
                    cooperations.Add(new
                    {
                        roleKey = cr["RoleKey"]?.ToString(),
                        userId = cr["UserId"]?.ToString(),
                        approverValue = cr["ApproverValue"]?.ToString(),
                        status = cr["Status"]?.ToString(),
                        action = cr["Action"]?.ToString(),
                        actedAtText = acted.HasValue ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc)) : null,
                        actorName = cr["ActorName"]?.ToString(),
                        approverDisplayText = cr["ApproverDisplayText"]?.ToString() ?? string.Empty
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
                createdAt = createdAt.HasValue ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(createdAt.Value, DateTimeKind.Utc)) : null,
                descriptorJson = descriptor,
                previewJson = preview,
                approvals,
                cooperations
            });
        }

        // ========== Private helpers ==========

        private sealed class ExcelRenderAreaInfo
        {
            public int UsedMaxRow { get; set; }
            public int UsedMaxCol { get; set; }
            public int RenderMaxRow { get; set; }
            public int RenderMaxCol { get; set; }
            public string UsedLastCellA1 { get; set; } = "A1";
            public string RenderLastCellA1 { get; set; } = "B2";
            public double WidthPx { get; set; }
            public double HeightPx { get; set; }
        }

        private sealed class ApprovalTableRowVm
        {
            public string RowType { get; set; } = string.Empty;
            public int GroupOrder { get; set; }
            public int SortOrder { get; set; }
            public string RoleKey { get; set; } = string.Empty;
            public string RoleText { get; set; } = string.Empty;
            public string DisplayName { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string ActedAtText { get; set; } = string.Empty;
        }
        private sealed class RenderDiagTailColVm
        {
            public string A1 { get; set; } = string.Empty;
            public int Col { get; set; }
            public bool Hidden { get; set; }
            public double Width { get; set; }
            public int UsedCount { get; set; }
            public List<string> Sample { get; set; } = new();
        }

        private sealed class RenderDiagTailRowVm
        {
            public int Row { get; set; }
            public bool Hidden { get; set; }
            public double Height { get; set; }
            public int UsedCount { get; set; }
            public List<string> Sample { get; set; } = new();
        }

        private static string BuildRenderAreaDiagJsonFromExcel(string excelPath, ILogger? log = null)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
                return "{}";

            try
            {
                using var wb = new XLWorkbook(excelPath);
                var ws = wb.Worksheets.First();

                var bbox = BuildVisibleBBox(ws);
                var usedMaxRow = Math.Max(1, bbox.usedMaxRow);
                var usedMaxCol = Math.Max(1, bbox.usedMaxCol);
                var renderMaxRow = Math.Min(XLHelper.MaxRowNumber, usedMaxRow + 1);
                var renderMaxCol = Math.Min(XLHelper.MaxColumnNumber, usedMaxCol + 1);

                var scanOptions = XLCellsUsedOptions.AllContents;

                int ExpandCol(IXLCell c)
                {
                    var row = c.Address.RowNumber;
                    var col = c.Address.ColumnNumber;

                    if (TryGetMergeBounds(bbox.merged, row, col, out var mb) && mb.r1 == row && mb.c1 == col)
                        return mb.c2;

                    return col;
                }

                int ExpandRow(IXLCell c)
                {
                    var row = c.Address.RowNumber;
                    var col = c.Address.ColumnNumber;

                    if (TryGetMergeBounds(bbox.merged, row, col, out var mb) && mb.r1 == row && mb.c1 == col)
                        return mb.r2;

                    return row;
                }

                var rowsNearEdge = new List<object>();
                foreach (var row in ws.RowsUsed(scanOptions))
                {
                    var lastCol = row.CellsUsed(scanOptions)
                        .Where(HasMeaningfulValue)
                        .Select(ExpandCol)
                        .DefaultIfEmpty(0)
                        .Max();

                    if (lastCol <= 0) continue;

                    rowsNearEdge.Add(new
                    {
                        row = row.RowNumber(),
                        lastCellA1 = ToA1(row.RowNumber(), lastCol),
                        lastCol
                    });
                }

                rowsNearEdge = rowsNearEdge
                    .OrderByDescending(x => (int)x.GetType().GetProperty("lastCol")!.GetValue(x)!)
                    .ThenByDescending(x => (int)x.GetType().GetProperty("row")!.GetValue(x)!)
                    .Take(8)
                    .Cast<object>()
                    .ToList();

                var mergeNearEdge = bbox.merged
                    .OrderByDescending(x => x.c2)
                    .ThenByDescending(x => x.r2)
                    .Take(8)
                    .Select(x => new
                    {
                        rangeA1 = ToA1Range(x.r1, x.c1, x.r2, x.c2),
                        row2 = x.r2,
                        col2 = x.c2
                    })
                    .Cast<object>()
                    .ToList();

                var tailCols = new List<RenderDiagTailColVm>();
                for (int c = Math.Max(1, usedMaxCol - 3); c <= Math.Min(XLHelper.MaxColumnNumber, usedMaxCol + 1); c++)
                {
                    var samples = ws.Column(c).CellsUsed(scanOptions)
                        .Where(HasMeaningfulValue)
                        .Take(5)
                        .Select(x => ToA1(ExpandRow(x), ExpandCol(x)))
                        .ToList();

                    tailCols.Add(new RenderDiagTailColVm
                    {
                        A1 = ToA1(1, c),
                        Col = c,
                        Hidden = ws.Column(c).IsHidden,
                        Width = Math.Round(ws.Column(c).Width <= 0 ? 8.43 : ws.Column(c).Width, 2),
                        UsedCount = ws.Column(c).CellsUsed(scanOptions).Count(HasMeaningfulValue),
                        Sample = samples
                    });
                }

                var tailRows = new List<RenderDiagTailRowVm>();
                for (int r = Math.Max(1, usedMaxRow - 3); r <= Math.Min(XLHelper.MaxRowNumber, usedMaxRow + 1); r++)
                {
                    var samples = ws.Row(r).CellsUsed(scanOptions)
                        .Where(HasMeaningfulValue)
                        .Take(5)
                        .Select(x => ToA1(ExpandRow(x), ExpandCol(x)))
                        .ToList();

                    tailRows.Add(new RenderDiagTailRowVm
                    {
                        Row = r,
                        Hidden = ws.Row(r).IsHidden,
                        Height = Math.Round(ws.Row(r).Height <= 0 ? 15 : ws.Row(r).Height, 2),
                        UsedCount = ws.Row(r).CellsUsed(scanOptions).Count(HasMeaningfulValue),
                        Sample = samples
                    });
                }

                var json = JsonSerializer.Serialize(new
                {
                    worksheet = ws.Name,
                    finalUsedLastCellA1 = ToA1(usedMaxRow, usedMaxCol),
                    finalUsedMaxCol = usedMaxCol,
                    finalUsedMaxRow = usedMaxRow,
                    renderLastCellA1 = ToA1(renderMaxRow, renderMaxCol),
                    renderMaxCol = renderMaxCol,
                    renderMaxRow = renderMaxRow,
                    rowsNearEdge,
                    mergeNearEdge,
                    tailCols,
                    tailRows
                });

                log?.LogInformation("DetailDX BuildRenderAreaDiagJson {Diag}", json);
                return json;
            }
            catch (Exception ex)
            {
                log?.LogWarning(ex, "BuildRenderAreaDiagJsonFromExcel failed path={Path}", excelPath);
                return "{}";
            }
        }

        private string BuildRoleText(string? roleKey, string rowType)
        {
            var order = ParseRoleOrder(roleKey);
            if (string.Equals(rowType, "drafter", StringComparison.OrdinalIgnoreCase)) return LocalizeOrFallback("DOC_Role_Submit", "상신");
            if (string.Equals(rowType, "approval", StringComparison.OrdinalIgnoreCase)) return LocalizeOrFallback("DOC_Role_ApprovalFmt", $"{order}차 결재", order);
            if (string.Equals(rowType, "cooperation", StringComparison.OrdinalIgnoreCase)) return LocalizeOrFallback("DOC_Role_CooperationFmt", $"협조 {order}", order);
            return roleKey ?? string.Empty;
        }

        private string LocalizeOrFallback(string key, string fallback, params object[] args)
        {
            try
            {
                var localized = _S[key, args];
                if (localized.ResourceNotFound) return fallback;
                var value = localized.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(value) || string.Equals(value, key, StringComparison.Ordinal)) return fallback;
                return value;
            }
            catch { return fallback; }
        }

        private static int ParseRoleOrder(string? roleKey)
        {
            if (string.IsNullOrWhiteSpace(roleKey)) return 9999;
            var m = Regex.Match(roleKey.Trim(), @"(\d+)$", RegexOptions.IgnoreCase);
            if (m.Success && int.TryParse(m.Groups[1].Value, out var n)) return n;
            return 9999;
        }

        private static string? TryGetJsonString(JsonElement parent, string name)
        {
            if (parent.ValueKind != JsonValueKind.Object) return null;
            if (!parent.TryGetProperty(name, out var el)) return null;
            return el.ValueKind switch
            {
                JsonValueKind.String => el.GetString(),
                JsonValueKind.Number => el.GetRawText(),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                _ => null
            };
        }

        private static string GetAnonymousString(object src, string propName)
        {
            var p = src.GetType().GetProperty(propName);
            if (p == null) return string.Empty;
            return Convert.ToString(p.GetValue(src)) ?? string.Empty;
        }

        private static int? GetAnonymousInt(object src, string propName)
        {
            var p = src.GetType().GetProperty(propName);
            if (p == null) return null;
            var v = p.GetValue(src);
            if (v == null) return null;
            if (v is int i) return i;
            return int.TryParse(Convert.ToString(v), out var n) ? n : (int?)null;
        }

        private string ResolveContentRootRelativePath(string? path)
        {
            var raw = (path ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            if (System.IO.Path.IsPathRooted(raw)) return raw;
            raw = raw.TrimStart('\\', '/').Replace('/', System.IO.Path.DirectorySeparatorChar).Replace('\\', System.IO.Path.DirectorySeparatorChar);
            return System.IO.Path.Combine(_env.ContentRootPath, raw);
        }

        private string MakeRelativeToContentRoot(string? absPath)
        {
            var raw = (absPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            try { if (System.IO.Path.IsPathRooted(raw)) return System.IO.Path.GetRelativePath(_env.ContentRootPath, raw); }
            catch { }
            return raw;
        }

        private async Task LogDocumentViewAsync(SqlConnection conn, string docId)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            var clientIp = HttpContext.Connection.RemoteIpAddress?.ToString() ?? string.Empty;
            var userAgent = Request.Headers["User-Agent"].ToString() ?? string.Empty;
            var tab = (Request?.Query["tab"].ToString() ?? string.Empty).Trim();
            var isSharedTab = string.Equals(tab, "shared", StringComparison.OrdinalIgnoreCase);

            var isSharedRecipient = false;
            try
            {
                await using var sc = conn.CreateCommand();
                sc.CommandText = @"SELECT TOP 1 1 FROM dbo.DocumentShares s WITH (NOLOCK) WHERE s.DocId=@DocId AND s.UserId=@UserId AND ISNULL(s.IsRevoked,0)=0 AND (s.ExpireAt IS NULL OR s.ExpireAt>SYSUTCDATETIME());";
                sc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                sc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                var o = await sc.ExecuteScalarAsync();
                isSharedRecipient = (o != null && o != DBNull.Value);
            }
            catch (Exception ex) { _log.LogWarning(ex, "DetailDX: shared recipient check failed. DocId={DocId}", docId); }

            string viewerRole = "Other";
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT CreatedBy,
       CASE WHEN EXISTS (SELECT 1 FROM dbo.DocumentApprovals da WHERE da.DocId=d.DocId AND da.UserId=@UserId) THEN 1 ELSE 0 END AS IsApprover
FROM dbo.Documents d WHERE d.DocId=@DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                await using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    var createdBy = r["CreatedBy"] as string;
                    var isApprover = Convert.ToInt32(r["IsApprover"]) == 1;
                    var isCreator = !string.IsNullOrWhiteSpace(createdBy) && string.Equals(createdBy, userId, StringComparison.OrdinalIgnoreCase);
                    if (isSharedTab) viewerRole = "Shared";
                    else if (isCreator && isApprover) viewerRole = "Creator+Approver";
                    else if (isCreator) viewerRole = "Creator";
                    else if (isApprover) viewerRole = "Approver";
                }
            }

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"INSERT INTO dbo.DocumentViewLogs (DocId,ViewerId,ViewerRole,ViewedAt,ClientIp,UserAgent) VALUES (@DocId,@ViewerId,@ViewerRole,SYSUTCDATETIME(),@ClientIp,@UserAgent);";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@ViewerId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                cmd.Parameters.Add(new SqlParameter("@ViewerRole", SqlDbType.NVarChar, 50) { Value = (object)viewerRole ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@ClientIp", SqlDbType.NVarChar, 64) { Value = (object)clientIp ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@UserAgent", SqlDbType.NVarChar, 512) { Value = (object)userAgent ?? DBNull.Value });
                await cmd.ExecuteNonQueryAsync();
            }

            if (isSharedTab && isSharedRecipient)
            {
                try
                {
                    await using var uc = conn.CreateCommand();
                    uc.CommandText = @"UPDATE dbo.DocumentShares SET IsRead=1 WHERE DocId=@DocId AND UserId=@UserId AND ISNULL(IsRevoked,0)=0 AND (ExpireAt IS NULL OR ExpireAt>SYSUTCDATETIME()) AND ISNULL(IsRead,0)=0;";
                    uc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    uc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                    await uc.ExecuteNonQueryAsync();
                }
                catch (Exception ex) { _log.LogWarning(ex, "DetailDX: shared read update failed. DocId={DocId}", docId); }
            }
        }

        private async Task<(bool canRecall, bool canApprove, bool canHold, bool canReject, bool canCooperate)> GetApprovalCapabilitiesAsync(string docId)
        {
            var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(userId))
                return (false, false, false, false, false);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            string? status = null;
            string? createdBy = null;

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"SELECT TOP 1 Status, CreatedBy FROM dbo.Documents WHERE DocId = @id;";
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                await using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    status = r["Status"] as string ?? string.Empty;
                    createdBy = r["CreatedBy"] as string ?? string.Empty;
                }
            }

            if (status == null) return (false, false, false, false, false);

            var st = status ?? string.Empty;
            var isCreator = string.Equals(createdBy ?? string.Empty, userId, StringComparison.OrdinalIgnoreCase);
            var isRecalled = st.Equals("Recalled", StringComparison.OrdinalIgnoreCase);
            var isFinal = st.StartsWith("Approved", StringComparison.OrdinalIgnoreCase)
                         || st.StartsWith("Rejected", StringComparison.OrdinalIgnoreCase);
            var canRecall = isCreator && st.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);

            bool canCooperate = false;
            await using (var coopCmd = conn.CreateCommand())
            {
                coopCmd.CommandText = @"SELECT TOP 1 1 FROM dbo.DocumentCooperations WHERE DocId=@id AND UserId=@uid AND ISNULL(Status,N'Pending')=N'Pending';";
                coopCmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                coopCmd.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = userId });
                var coopObj = await coopCmd.ExecuteScalarAsync();
                canCooperate = (coopObj != null && coopObj != DBNull.Value);
            }

            int currentStep = 0;
            await using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = @"SELECT MIN(StepOrder) FROM dbo.DocumentApprovals WHERE DocId=@id AND ISNULL(Status,N'Pending') LIKE N'Pending%';";
                cmd2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                var obj = await cmd2.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value) currentStep = Convert.ToInt32(obj);
            }

            if (currentStep == 0 || isFinal || isRecalled)
                return (canRecall, false, false, false, canCooperate);

            string? stepUserId = null;
            string stepStatus = string.Empty;

            await using (var cmd3 = conn.CreateCommand())
            {
                cmd3.CommandText = @"SELECT TOP 1 UserId, ISNULL(Status,N'Pending') AS Status FROM dbo.DocumentApprovals WHERE DocId=@id AND StepOrder=@step;";
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
                return (canRecall, false, false, false, canCooperate);

            var isMyStep = string.Equals(stepUserId, userId, StringComparison.OrdinalIgnoreCase);
            var isPending = stepStatus.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);
            var canDo = isMyStep && isPending && !isRecalled && !isFinal;

            return (canRecall, canDo, canDo, canDo, canCooperate);
        }

        private async Task<string?> GetDocumentAuthorUserIdAsync(SqlConnection conn, string docId)
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT TOP (1) NULLIF(LTRIM(RTRIM(CreatedBy)),N'') AS CreatedBy FROM dbo.Documents WHERE DocId=@DocId;";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            var obj = await cmd.ExecuteScalarAsync();
            var s = (obj == null || obj == DBNull.Value) ? null : (obj.ToString() ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(s) ? null : s;
        }

        private async Task<List<string>> GetNextApproverUserIdsFromDbAsync(SqlConnection conn, string docId, int nextStep)
        {
            var list = new List<string>();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT DISTINCT
       NULLIF(LTRIM(RTRIM(a.UserId)),N'')        AS UserId,
       NULLIF(LTRIM(RTRIM(a.ApproverValue)),N'') AS ApproverValue
FROM dbo.DocumentApprovals a
WHERE a.DocId=@DocId AND a.StepOrder=@Step AND ISNULL(a.Status,N'Pending') LIKE N'Pending%';";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = nextStep });

            var approverValuesToResolve = new List<string>();
            await using (var r = await cmd.ExecuteReaderAsync())
            {
                while (await r.ReadAsync())
                {
                    var uid = r["UserId"] as string;
                    var av = r["ApproverValue"] as string;
                    if (!string.IsNullOrWhiteSpace(uid)) { list.Add(uid.Trim()); continue; }
                    if (!string.IsNullOrWhiteSpace(av)) approverValuesToResolve.Add(av.Trim());
                }
            }

            var toBackfill = new List<(string approverValue, string resolvedUserId)>();
            foreach (var av in approverValuesToResolve.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var resolved = await TryResolveUserIdAsync(conn, tx: null, value: av);
                if (!string.IsNullOrWhiteSpace(resolved)) { list.Add(resolved.Trim()); toBackfill.Add((av, resolved.Trim())); }
            }

            foreach (var x in toBackfill
                .Select(t => new { approverValue = t.approverValue.Trim(), resolvedUserId = t.resolvedUserId.Trim(), key = t.approverValue.Trim() + "\u001F" + t.resolvedUserId.Trim() })
                .Where(x => x.approverValue.Length > 0 && x.resolvedUserId.Length > 0)
                .DistinctBy(x => x.key, StringComparer.OrdinalIgnoreCase))
            {
                await using var up = conn.CreateCommand();
                up.CommandText = @"UPDATE dbo.DocumentApprovals SET UserId=@UserId WHERE DocId=@DocId AND StepOrder=@Step AND (UserId IS NULL OR LTRIM(RTRIM(UserId))=N'') AND ApproverValue=@ApproverValue;";
                up.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = x.resolvedUserId });
                up.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                up.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = nextStep });
                up.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 256) { Value = x.approverValue });
                await up.ExecuteNonQueryAsync();
            }

            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static async Task<string?> TryResolveUserIdAsync(SqlConnection conn, SqlTransaction? tx, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            await using var cmd = conn.CreateCommand();
            if (tx != null) cmd.Transaction = tx;
            cmd.CommandText = @"SELECT TOP 1 Id FROM dbo.AspNetUsers WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);";
            cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            var id = (string?)await cmd.ExecuteScalarAsync();
            if (!string.IsNullOrWhiteSpace(id)) return id;

            await using var cmd2 = conn.CreateCommand();
            if (tx != null) cmd2.Transaction = tx;
            cmd2.CommandText = @"SELECT TOP 1 COALESCE(p.UserId,u.Id) FROM dbo.UserProfiles p LEFT JOIN dbo.AspNetUsers u ON u.Id=p.UserId WHERE p.UserId=@v OR p.Email=@v OR p.DisplayName=@v OR p.Name=@v;";
            cmd2.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            id = (string?)await cmd2.ExecuteScalarAsync();
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        private static async Task<List<string>> GetNextApproverEmailsFromDbAsync(SqlConnection conn, string docId, int stepOrder)
        {
            var result = new List<string>();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT ApproverValue FROM dbo.DocumentApprovals
WHERE DocId=@DocId AND StepOrder=@StepOrder AND ISNULL(Status,N'Pending') LIKE N'Pending%'
  AND ApproverValue IS NOT NULL AND LTRIM(RTRIM(ApproverValue))<>N'';";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = stepOrder });
            await using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                var v = r[0] as string;
                if (!string.IsNullOrWhiteSpace(v)) result.Add(v.Trim());
            }
            return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // ========== DTOs ==========
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
            public List<CommentFileDto>? files { get; set; }
        }

        public sealed class CommentFileDto
        {
            public string? FileKey { get; set; }
            public string? OriginalName { get; set; }
            public string? ContentType { get; set; }
            public long ByteSize { get; set; }
        }

        public sealed class UpdateSharesDto
        {
            public string? DocId { get; set; }
            public List<string>? SelectedRecipientUserIds { get; set; }
        }
    }
}
