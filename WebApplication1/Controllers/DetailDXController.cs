// 2026.03.17 Added: DetailDX 구성
// 2026.03.18 Changed: DxTemp 임시파일 생성 제거 — ReadOnly 조회이므로 원본 파일 직접 사용
// 2026.03.23 Added: Cooperate 기능 추가, DetailController 전체 기능 마이그레이션
// 2026.03.24 Added: 협조 스탬프 지원 — DescriptorJson.Cooperations에서 CellA1 파싱, UserProfiles.SignatureRelativePath JOIN
// 2026.03.25 Fixed: 협조 셀맵 파싱 시 cellA1 키 추가 탐색 (Cell.A1 / A1 / cellA1 / CellA1 순서로 탐색)
// 2026.03.26 Fixed: Documents.DescriptorJson에 cellA1 없을 때 DocTemplateVersion.DescriptorJson Cooperations[].Cell.A1 fallback 추가
using DevExpress.AspNetCore.Spreadsheet;
using DevExpress.Spreadsheet;
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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Claims;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WebApplication1.Models;
using WebApplication1.Services;
using GdiBitmap = System.Drawing.Bitmap;
using GdiColor = System.Drawing.Color;
using GdiImage = System.Drawing.Image;
using GdiPixelFormat = System.Drawing.Imaging.PixelFormat;
using GdiRectangle = System.Drawing.Rectangle;
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
        // 2026.06.15 Added: DetailDX warm-up workbook 최초 생성 시 동시 요청 충돌을 방지한다.
        private static readonly object _detailDxWarmupWorkbookLock = new();

        // 2026.06.23 Added: DetailDX 스탬프 작업 파일 및 서버 삽입 처리 기준 상수 추가 Contents ComposeDX 세션 저장 방식과 동일하게 Spreadsheet 세션에서 이미지 삽입 후 저장한다
        private static readonly object _detailDxStampFileLock = new();
        // 2026.07.02 Changed: 스탬프 규칙 버전 갱신 Contents 여백 제거와 셀 비율 90퍼센트 배치 방식으로 기존 삽입분을 다시 처리한다
        private const string StampRuleVersion = "DetailDXDevExpressStampV25";
        private const string ApprovalLineType = "Approval";
        private const string CooperationLineType = "Cooperation";
        private const string StampPicturePrefix = "EB_STAMP_";

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

            var currentCulture = DocControllerHelper.NormalizeCultureName(CultureInfo.CurrentUICulture.Name);

            string? templateCode = null;
            string? templateTitle = null;
            string? status = null;
            string? descriptorJson = null;
            string? outputPath = null;
            string? compCd = null;
            string? compTimeZoneId = null;
            string? creatorNameRaw = null;
            DateTime? createdAtUtc = null;
            long? documentTemplateVersionId = null;

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
       d.CreatedByName,
       d.CreatedAt,
       d.TemplateVersionId,
       COALESCE(NULLIF(LTRIM(RTRIM(cm.TimeZoneId)), N''), N'Asia/Seoul') AS CompTimeZoneId
FROM dbo.Documents d
LEFT JOIN dbo.CompMasters cm ON cm.CompCd = d.CompCd
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
                compTimeZoneId = rd["CompTimeZoneId"]?.ToString();
                creatorNameRaw = rd["CreatedByName"] as string ?? string.Empty;
                createdAtUtc = rd["CreatedAt"] is DateTime cdt ? cdt : (DateTime?)null;
                documentTemplateVersionId = rd["TemplateVersionId"] == DBNull.Value
                    ? null
                    : Convert.ToInt64(rd["TemplateVersionId"], CultureInfo.InvariantCulture);
            }

            string createdAtText = FormatDetailLocalMinute(createdAtUtc, compTimeZoneId, currentCulture);

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
                    recalledAtText = FormatDetailLocalMinute(dt, compTimeZoneId, currentCulture);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: recalled time load failed for {DocId}", id);
            }

            long? templateVersionId = null;
            string templateVisualRangeA1 = string.Empty;
            string templateVisualSource = string.Empty;
            int templateVisualWidthPx = 0;
            int templateVisualHeightPx = 0;

            if (documentTemplateVersionId.HasValue && documentTemplateVersionId.Value > 0)
            {
                await using var tvCmd = conn.CreateCommand();
                tvCmd.CommandText = @"
SELECT TOP 1
       v.Id,
       ISNULL(v.VisualRangeA1, N'') AS VisualRangeA1,
       ISNULL(v.VisualSource,  N'') AS VisualSource,
       ISNULL(v.VisualWidthPx,  0) AS VisualWidthPx,
       ISNULL(v.VisualHeightPx, 0) AS VisualHeightPx
FROM dbo.DocTemplateVersion v
WHERE v.Id = @TemplateVersionId;";
                tvCmd.Parameters.Add(new SqlParameter("@TemplateVersionId", SqlDbType.BigInt) { Value = documentTemplateVersionId.Value });

                await using var tvRd = await tvCmd.ExecuteReaderAsync();
                if (await tvRd.ReadAsync())
                {
                    templateVersionId = tvRd.IsDBNull(0) ? null : Convert.ToInt64(tvRd.GetValue(0));
                    templateVisualRangeA1 = tvRd["VisualRangeA1"]?.ToString() ?? string.Empty;
                    templateVisualSource = tvRd["VisualSource"]?.ToString() ?? string.Empty;
                    templateVisualWidthPx = tvRd["VisualWidthPx"] == DBNull.Value ? 0 : Convert.ToInt32(tvRd["VisualWidthPx"]);
                    templateVisualHeightPx = tvRd["VisualHeightPx"] == DBNull.Value ? 0 : Convert.ToInt32(tvRd["VisualHeightPx"]);
                }
            }

            // 2026.06.15 Changed: 문서에 저장된 TemplateVersionId가 없을 때만 기존 TemplateCode 최신 버전 조회를 fallback으로 사용한다.
            if (!templateVersionId.HasValue && !string.IsNullOrWhiteSpace(templateCode))
            {
                await using var tvCmd = conn.CreateCommand();
                tvCmd.CommandText = @"
SELECT TOP 1
       v.Id,
       ISNULL(v.VisualRangeA1, N'') AS VisualRangeA1,
       ISNULL(v.VisualSource,  N'') AS VisualSource,
       ISNULL(v.VisualWidthPx,  0) AS VisualWidthPx,
       ISNULL(v.VisualHeightPx, 0) AS VisualHeightPx
FROM dbo.DocTemplateVersion v
INNER JOIN dbo.DocTemplateMaster m ON v.TemplateId = m.Id
WHERE m.DocCode = @TemplateCode
ORDER BY v.VersionNo DESC, v.Id DESC;";
                tvCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode! });

                await using var tvRd = await tvCmd.ExecuteReaderAsync();
                if (await tvRd.ReadAsync())
                {
                    templateVersionId = tvRd.IsDBNull(0) ? null : Convert.ToInt64(tvRd.GetValue(0));
                    templateVisualRangeA1 = tvRd["VisualRangeA1"]?.ToString() ?? string.Empty;
                    templateVisualSource = tvRd["VisualSource"]?.ToString() ?? string.Empty;
                    templateVisualWidthPx = tvRd["VisualWidthPx"] == DBNull.Value ? 0 : Convert.ToInt32(tvRd["VisualWidthPx"]);
                    templateVisualHeightPx = tvRd["VisualHeightPx"] == DBNull.Value ? 0 : Convert.ToInt32(tvRd["VisualHeightPx"]);
                }
            }

            // 2026.06.23 Changed: DetailDX 작업 파일과 DevExpress documentId를 같은 세션 단위로 생성 Contents documentId 재사용으로 기존 메모리 문서가 열리는 문제 방지
            var dxSourceAbsPath = ResolveContentRootRelativePath(outputPath) ?? string.Empty;
            var detailDxStampSessionId = DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture)
                + "_"
                + Guid.NewGuid().ToString("N");

            var dxDocumentId = "detail_"
                + SafeStampFilePart(id)
                + "_"
                + SafeStampFilePart(detailDxStampSessionId);

            var dxOpenAbsPath = CreateDetailDxStampWorkbookCopy(id!, dxSourceAbsPath, detailDxStampSessionId);
            var stampBackfillRequired = await HasPendingStampRowsAsync(conn, id!, templateVersionId, descriptorJson);

            _log.LogInformation(
                "DetailDX: dxSourceAbsPath={SourcePath} dxOpenAbsPath={OpenPath} exists={Exists} templateVersionId={TemplateVersionId} visualRange={VisualRangeA1} visualWidth={VisualWidthPx} visualHeight={VisualHeightPx} visualSource={VisualSource} stampBackfillRequired={StampBackfillRequired}",
                dxSourceAbsPath,
                dxOpenAbsPath,
                System.IO.File.Exists(dxOpenAbsPath ?? ""),
                templateVersionId,
                templateVisualRangeA1,
                templateVisualWidthPx,
                templateVisualHeightPx,
                templateVisualSource,
                stampBackfillRequired);

            // 2026.06.12 Changed: DetailDX는 ComposeDX와 동일하게 DB 저장값만 사용한다.
            // output xlsx 재스캔/Preview 재생성/RenderDiag 계산은 하지 않는다.
            var previewJson = "{}";
            if (templateVersionId.HasValue)
            {
                try
                {
                    await using var tvFbCmd = conn.CreateCommand();
                    tvFbCmd.CommandText = @"
SELECT TOP 1 PreviewJson
FROM dbo.DocTemplateVersion
WHERE Id = @VersionId;";
                    tvFbCmd.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = templateVersionId.Value });
                    var tplPreviewObj = await tvFbCmd.ExecuteScalarAsync();
                    var tplPreview = tplPreviewObj == null || tplPreviewObj == DBNull.Value ? null : tplPreviewObj.ToString();
                    if (!string.IsNullOrWhiteSpace(tplPreview) && tplPreview != "{}")
                        previewJson = tplPreview;
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "DetailDX: template previewJson load failed for {DocId}", id);
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
        a.ApproverDisplayText,
        a.SignaturePath
FROM dbo.DocumentApprovals a
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
                        actedAtText = FormatDetailLocalMinute(acted, compTimeZoneId, currentCulture),
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
                ac2.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = templateVersionId.Value });

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
                        uploadedAtText = FormatDetailLocalMinute(upUtc, compTimeZoneId, currentCulture);
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

            var cooperationCellMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var cooperationRoleCellMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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

                                var av =
                                    TryGetJsonString(coop, "ApproverValue")
                                    ?? TryGetJsonString(coop, "approverValue")
                                    ?? TryGetJsonString(coop, "value")
                                    ?? string.Empty;

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
            DisplayName = creatorNameRaw ?? string.Empty,
            Status      = string.Equals(status, "Recalled", StringComparison.OrdinalIgnoreCase) ? "Recalled" : "Created",
            ActedAtText = string.Equals(status, "Recalled", StringComparison.OrdinalIgnoreCase) ? recalledAtText : createdAtText
        }
    };

            foreach (var a in approvals)
            {
                var roleKey = GetAnonymousString(a, "roleKey");
                var approverDisplayText = GetAnonymousString(a, "approverDisplayText");

                approvalTableRows.Add(new ApprovalTableRowVm
                {
                    RowType = "approval",
                    GroupOrder = 1,
                    SortOrder = GetAnonymousInt(a, "step") ?? ParseRoleOrder(roleKey),
                    RoleKey = roleKey,
                    RoleText = BuildRoleText(roleKey, "approval"),
                    DisplayName = approverDisplayText ?? string.Empty,
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
       dc.ApproverDisplayText,
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
                        actedAtText = FormatDetailLocalMinute(acted, compTimeZoneId, currentCulture),
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
                        DisplayName = approverDisplayText ?? string.Empty,
                        Status = cr["Status"]?.ToString() ?? string.Empty,
                        ActedAtText = FormatDetailLocalMinute(acted, compTimeZoneId, currentCulture)
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
                    var viewedAtText = FormatDetailLocalMinute(viewedUtc, compTimeZoneId, currentCulture);
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

            // 2026.06.12 Changed: DetailDX 렌더 크기는 DB 저장값을 그대로 사용한다.
            // 여기서는 xlsx를 다시 열어 폭/높이를 계산하지 않는다.
            var renderUsedLastCellA1 = templateVisualRangeA1;
            var renderLastCellA1 = templateVisualRangeA1;
            var renderMaxRow = 0;
            var renderMaxCol = 0;

            if (TryParseVisualRangeLastCell(templateVisualRangeA1, out var dbLastRow, out var dbLastCol, out var dbLastCellA1))
            {
                renderUsedLastCellA1 = dbLastCellA1;
                renderLastCellA1 = dbLastCellA1;
                renderMaxRow = dbLastRow;
                renderMaxCol = dbLastCol;
            }

            ViewBag.RenderDiagJson = "{}";
            ViewBag.RenderUsedLastCellA1 = string.IsNullOrWhiteSpace(renderUsedLastCellA1) ? "A1" : renderUsedLastCellA1;
            ViewBag.RenderLastCellA1 = string.IsNullOrWhiteSpace(renderLastCellA1) ? "A1" : renderLastCellA1;
            ViewBag.RenderMaxRow = renderMaxRow;
            ViewBag.RenderMaxCol = renderMaxCol;
            ViewBag.RenderWidthPx = templateVisualWidthPx;
            ViewBag.RenderHeightPx = templateVisualHeightPx;


            var caps = await GetApprovalCapabilitiesAsync(id!);

            ViewData["Title"] = _S["DOC_Title_Detail"];
            ViewData["DisableDxAll"] = true;

            ViewBag.DocId = id;
            ViewBag.DocumentId = id;
            ViewBag.TemplateCode = templateCode ?? string.Empty;
            ViewBag.TemplateTitle = templateTitle ?? string.Empty;
            ViewBag.Status = status ?? string.Empty;
            ViewBag.CompCd = compCd ?? string.Empty;
            ViewBag.CreatorName = creatorNameRaw ?? string.Empty;
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
            ViewBag.SourceExcelPath = dxSourceAbsPath ?? string.Empty;
            ViewBag.StampWorkPath = dxOpenAbsPath ?? string.Empty;
            ViewBag.StampBackfillRequired = stampBackfillRequired;

            var http = HttpContext;
            if (http == null)
                throw new InvalidOperationException("HttpContext is not available.");
            var request = http.Request;

            ViewBag.DxCallbackUrl = "/Doc/dx-callback";
            ViewBag.DxDocumentId = dxDocumentId;
            ViewBag.StampSessionId = detailDxStampSessionId;
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

        // 2026.06.12 Added: DetailDX 기본 렌더 크기는 템플릿 저장 시 확정한 DB VisualMetric을 사용한다.
        private static ExcelRenderAreaInfo BuildRenderAreaInfoFromTemplateVisualMetrics(string? visualRangeA1, int visualWidthPx, int visualHeightPx)
        {
            var info = new ExcelRenderAreaInfo
            {
                UsedMaxRow = 1,
                UsedMaxCol = 1,
                RenderMaxRow = 2,
                RenderMaxCol = 2,
                UsedLastCellA1 = "A1",
                RenderLastCellA1 = "B2",
                WidthPx = Math.Max(0, visualWidthPx),
                HeightPx = Math.Max(0, visualHeightPx)
            };

            if (TryParseVisualRangeLastCell(visualRangeA1, out var lastRow, out var lastCol, out var lastCellA1))
            {
                info.UsedMaxRow = Math.Max(1, lastRow);
                info.UsedMaxCol = Math.Max(1, lastCol);
                info.RenderMaxRow = Math.Max(2, lastRow);
                info.RenderMaxCol = Math.Max(2, lastCol);
                info.UsedLastCellA1 = lastCellA1;
                info.RenderLastCellA1 = lastCellA1;
            }

            return info;
        }

        private static bool TryParseVisualRangeLastCell(string? visualRangeA1, out int lastRow, out int lastCol, out string lastCellA1)
        {
            lastRow = 0;
            lastCol = 0;
            lastCellA1 = "A1";

            var s = (visualRangeA1 ?? string.Empty).Trim().ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(s)) return false;

            var end = s.Contains(':') ? s[(s.LastIndexOf(':') + 1)..] : s;
            var m = Regex.Match(end, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
            if (!m.Success) return false;

            int col = 0;
            foreach (var ch in m.Groups[1].Value.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return false;
                col = col * 26 + (ch - 'A' + 1);
            }

            if (!int.TryParse(m.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row))
                return false;

            if (row <= 0 || col <= 0) return false;

            lastRow = row;
            lastCol = col;
            lastCellA1 = ToA1(row, col);
            return true;
        }

        // 2026.06.23 Added: DB VisualRangeA1 마지막 셀 주소 변환 헬퍼 복원 Contents ClosedXML 렌더 재계산 코드 제거 중 남은 ToA1 호출을 처리한다
        private static string ToA1(int row1, int col1)
        {
            if (row1 <= 0 || col1 <= 0)
                return string.Empty;

            var letters = string.Empty;
            var n = col1;

            while (n > 0)
            {
                n--;
                letters = (char)('A' + (n % 26)) + letters;
                n /= 26;
            }

            return letters + row1.ToString(CultureInfo.InvariantCulture);
        }

        // 2026.06.23 Removed: xlsx 재스캔 기반 렌더 범위 재계산 함수 제거 Contents DetailDX는 DB VisualMetric만 사용한다

        // ========== ApproveOrHold ==========
        [HttpPost("ApproveOrHoldDX")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        // 2026.05.21 Changed 승인 후 문서 상태를 실제 남은 대기 결재 차수 기준으로 갱신하고 중복 상태 업데이트와 중복 알림 블록 제거
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
                var requesterId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

                var csR = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(csR);
                await conn.OpenAsync();

                // 1) 작성자 본인만 회수 가능
                var authorId = await GetDocumentAuthorUserIdAsync(conn, dto.docId!);
                if (string.IsNullOrWhiteSpace(authorId) ||
                    string.IsNullOrWhiteSpace(requesterId) ||
                    !string.Equals(authorId.Trim(), requesterId.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return Forbid();
                }

                await using var tx = conn.BeginTransaction();

                try
                {
                    string docStatus = string.Empty;
                    await using (var ds = conn.CreateCommand())
                    {
                        ds.Transaction = tx;
                        ds.CommandText = @"
SELECT TOP (1) Status
FROM dbo.Documents
WHERE DocId = @DocId;";
                        ds.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId! });
                        docStatus = (await ds.ExecuteScalarAsync() as string) ?? string.Empty;
                    }

                    // 2) 문서 상태가 Pending 계열이 아니면 회수 불가
                    if (string.IsNullOrWhiteSpace(docStatus) ||
                        !docStatus.StartsWith("Pending", StringComparison.OrdinalIgnoreCase))
                    {
                        await tx.RollbackAsync();
                        return Conflict(new { messages = new[] { "DOC_Err_BadRequest" } });
                    }

                    // 3) 결재 진행 이력이 1건이라도 있으면 회수 불가
                    int progressedApprovalCount = 0;
                    await using (var ckAp = conn.CreateCommand())
                    {
                        ckAp.Transaction = tx;
                        ckAp.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND (
        ActedAt IS NOT NULL
        OR (
            NULLIF(LTRIM(RTRIM(ISNULL(Action, N''))), N'') IS NOT NULL
            AND ISNULL(Action, N'') <> N'Recalled'
        )
      );";
                        ckAp.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId! });
                        progressedApprovalCount = Convert.ToInt32(await ckAp.ExecuteScalarAsync());
                    }

                    if (progressedApprovalCount > 0)
                    {
                        await tx.RollbackAsync();
                        return Conflict(new { messages = new[] { "DOC_Err_BadRequest" } });
                    }

                    // 4) 협조 진행 이력이 1건이라도 있으면 회수 불가
                    int progressedCooperationCount = 0;
                    await using (var ckCo = conn.CreateCommand())
                    {
                        ckCo.Transaction = tx;
                        ckCo.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentCooperations
WHERE DocId = @DocId
  AND (
        ActedAt IS NOT NULL
        OR (
            NULLIF(LTRIM(RTRIM(ISNULL(Action, N''))), N'') IS NOT NULL
            AND ISNULL(Action, N'') <> N'Recalled'
        )
      );";
                        ckCo.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId! });
                        progressedCooperationCount = Convert.ToInt32(await ckCo.ExecuteScalarAsync());
                    }

                    if (progressedCooperationCount > 0)
                    {
                        await tx.RollbackAsync();
                        return Conflict(new { messages = new[] { "DOC_Err_BadRequest" } });
                    }

                    // 5) 문서 상태만 회수 처리
                    int docAffected = 0;
                    await using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
UPDATE dbo.Documents
SET Status = N'Recalled'
WHERE DocId = @DocId
  AND ISNULL(Status, N'') LIKE N'Pending%';";
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId! });
                        docAffected = await cmd.ExecuteNonQueryAsync();
                    }

                    if (docAffected <= 0)
                    {
                        await tx.RollbackAsync();
                        return Conflict(new { messages = new[] { "DOC_Err_BadRequest" } });
                    }

                    await tx.CommitAsync();
                    return Json(new { ok = true, docId = dto.docId, status = "Recalled" });
                }
                catch (Exception exRecall)
                {
                    try { await tx.RollbackAsync(); } catch { }
                    _log.LogError(exRecall, "Recall failed docId={docId}", dto.docId);
                    return StatusCode(StatusCodes.Status500InternalServerError,
                        new { messages = new[] { "DOC_Err_RequestFailed" } });
                }
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
       u.SignatureRelativePath         AS SignaturePath
FROM dbo.UserProfiles AS u
LEFT JOIN dbo.PositionMasters     AS pm ON pm.CompCd = u.CompCd AND pm.Id = u.PositionId
LEFT JOIN dbo.PositionMasterLoc  AS pl ON pl.PositionId = pm.Id AND pl.LangCode = N'ko'
WHERE u.UserId = @uid;";
                    up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });
                    await using var r = await up.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        var disp = r["DisplayName"] as string ?? string.Empty;
                        var pos = r["PositionName"] as string ?? string.Empty;
                        var sig = r["SignaturePath"] as string ?? string.Empty;

                        var actorDisplayText = string.Join(" ", new[] { pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));
                        if (!string.IsNullOrWhiteSpace(actorDisplayText))
                            approverName = actorDisplayText;

                        coopDisplayText = actorDisplayText;
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

                return await ApplyStampsAndReturnAsync(dto, "Cooperated");
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
       COALESCE(pl.Name, pm.Name, N'') AS PositionName
FROM dbo.UserProfiles AS u
LEFT JOIN dbo.PositionMasters   AS pm ON pm.CompCd = u.CompCd AND pm.Id = u.PositionId
LEFT JOIN dbo.PositionMasterLoc AS pl ON pl.PositionId = pm.Id AND pl.LangCode = N'ko'
WHERE u.UserId = @uid;";
                    up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });

                    await using var r = await up.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        var disp = r["DisplayName"] as string ?? string.Empty;
                        var pos = r["PositionName"] as string ?? string.Empty;

                        var actorDisplayText = string.Join(" ", new[] { pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));
                        if (!string.IsNullOrWhiteSpace(actorDisplayText))
                            approverName = actorDisplayText;

                        coopDisplayText = actorDisplayText;
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
       u.SignatureRelativePath         AS SignaturePath
FROM dbo.UserProfiles AS u
LEFT JOIN dbo.PositionMasters   AS pm ON pm.CompCd = u.CompCd AND pm.Id = u.PositionId
LEFT JOIN dbo.PositionMasterLoc AS pl ON pl.PositionId = pm.Id AND pl.LangCode = N'ko'
WHERE u.UserId = @uid;";
                up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId2 });

                await using var r = await up.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    var disp = r["DisplayName"] as string ?? string.Empty;
                    var pos = r["PositionName"] as string ?? string.Empty;
                    var sig = r["SignaturePath"] as string ?? string.Empty;

                    var actorDisplayText = string.Join(" ", new[] { pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));
                    if (!string.IsNullOrWhiteSpace(actorDisplayText))
                        approverName2 = actorDisplayText;

                    approverDisplayText = actorDisplayText;
                    signatureRelativePath = string.IsNullOrWhiteSpace(sig) ? null : sig;
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "approve/hold/reject: profile snapshot failed docId={docId}", dto.docId);
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
                // 문서의 현재 결재 차수 = 최소 Pending StepOrder
                await using (var findCurrentStep = conn.CreateCommand())
                {
                    findCurrentStep.CommandText = @"
SELECT MIN(StepOrder)
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND ISNULL(Status, N'Pending') LIKE N'Pending%';";
                    findCurrentStep.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    var stepObj = await findCurrentStep.ExecuteScalarAsync();
                    if (stepObj != null && stepObj != DBNull.Value)
                        currentStep = Convert.ToInt32(stepObj);
                }

                if (currentStep == null || currentStep.Value <= 0)
                    return Forbid();

                // 현재 최소 Pending 차수에 대해서만 본인 결재 가능
                bool isMyCurrentStep = false;
                await using (var mineStep = conn.CreateCommand())
                {
                    mineStep.CommandText = @"
SELECT TOP (1) 1
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder = @Step
  AND ISNULL(Status, N'Pending') LIKE N'Pending%'
  AND (
        NULLIF(LTRIM(RTRIM(UserId)), N'') = @UserId
     OR NULLIF(LTRIM(RTRIM(ApproverValue)), N'') = @UserId
  );";
                    mineStep.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    mineStep.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = currentStep.Value });
                    mineStep.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId2 });

                    var mineObj = await mineStep.ExecuteScalarAsync();
                    isMyCurrentStep = (mineObj != null && mineObj != DBNull.Value);
                }

                if (!isMyCurrentStep)
                    return Forbid();

                var step = currentStep.Value;
                actedStep = step;

                if (actionLower == "approve")
                {
                    bool isCurrentFinalApprovalStep = true;
                    await using (var nextStepChk = conn.CreateCommand())
                    {
                        nextStepChk.CommandText = @"
SELECT TOP 1 1
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder > @Step
ORDER BY StepOrder;";
                        nextStepChk.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId! });
                        nextStepChk.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = currentStep.Value });

                        var nextObj = await nextStepChk.ExecuteScalarAsync();
                        isCurrentFinalApprovalStep = (nextObj == null || nextObj == DBNull.Value);
                    }

                    if (isCurrentFinalApprovalStep)
                    {
                        int remainingCooperationCount = 0;
                        await using (var remainCoopCmd = conn.CreateCommand())
                        {
                            remainCoopCmd.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentCooperations
WHERE DocId = @DocId
  AND ISNULL(Status, N'Pending') NOT IN (N'Cooperated', N'Recalled');";
                            remainCoopCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId! });

                            remainingCooperationCount = Convert.ToInt32(await remainCoopCmd.ExecuteScalarAsync());
                        }

                        if (remainingCooperationCount > 0)
                            return Conflict(new { messages = new[] { "DOC_Err_BadRequest" } });
                    }
                }

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

                if (actionLower == "approve")
                {
                    int? nextPendingStep = null;

                    await using (var nextPendingCmd = conn.CreateCommand())
                    {
                        nextPendingCmd.CommandText = @"
SELECT MIN(StepOrder)
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder > @Step
  AND ISNULL(Status, N'Pending') = N'Pending';";
                        nextPendingCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        nextPendingCmd.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = step });

                        var nextPendingObj = await nextPendingCmd.ExecuteScalarAsync();
                        if (nextPendingObj != null && nextPendingObj != DBNull.Value)
                            nextPendingStep = Convert.ToInt32(nextPendingObj);
                    }

                    if (nextPendingStep == null)
                    {
                        newStatus = "Approved";

                        await using (var upDocApproved = conn.CreateCommand())
                        {
                            upDocApproved.CommandText = @"
UPDATE dbo.Documents
SET Status = N'Approved'
WHERE DocId = @DocId;";
                            upDocApproved.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            await upDocApproved.ExecuteNonQueryAsync();
                        }
                    }
                    else
                    {
                        newStatus = $"PendingA{nextPendingStep.Value}";
                        nextStepForNotify = nextPendingStep.Value;

                        await using (var upDocPending = conn.CreateCommand())
                        {
                            upDocPending.CommandText = @"
UPDATE dbo.Documents
SET Status = @Status
WHERE DocId = @DocId;";
                            upDocPending.Parameters.Add(new SqlParameter("@Status", SqlDbType.NVarChar, 20) { Value = newStatus });
                            upDocPending.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            await upDocPending.ExecuteNonQueryAsync();
                        }

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
                            chkLast.Parameters.Add(new SqlParameter("@Next", SqlDbType.Int) { Value = nextPendingStep.Value });

                            var stepAfterNextObj = await chkLast.ExecuteScalarAsync();
                            if (stepAfterNextObj != null && stepAfterNextObj != DBNull.Value)
                                stepAfterNext = Convert.ToInt32(stepAfterNextObj);
                        }

                        var shouldNotifyNextApprover = true;
                        var isNextFinalStep = stepAfterNext == null;

                        if (isNextFinalStep)
                        {
                            int coopPending = 0;

                            await using (var chkCoop = conn.CreateCommand())
                            {
                                chkCoop.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentCooperations
WHERE DocId = @DocId
  AND ISNULL(Status, N'Pending') NOT IN (N'Cooperated', N'Recalled');";
                                chkCoop.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });

                                coopPending = Convert.ToInt32(await chkCoop.ExecuteScalarAsync());
                            }

                            shouldNotifyNextApprover = coopPending == 0;
                        }

                        if (shouldNotifyNextApprover)
                        {
                            try
                            {
                                var nextIds = await GetNextApproverUserIdsFromDbAsync(conn, dto.docId!, nextPendingStep.Value);
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
                else
                {
                    newStatus = actionLower switch
                    {
                        "hold" => $"PendingHoldA{step}",
                        "reject" => $"RejectedA{step}",
                        _ => "Updated"
                    };

                    await using (var upDocStatus = conn.CreateCommand())
                    {
                        upDocStatus.CommandText = @"
UPDATE dbo.Documents
SET Status = @Status
WHERE DocId = @DocId;";
                        upDocStatus.Parameters.Add(new SqlParameter("@Status", SqlDbType.NVarChar, 20) { Value = newStatus });
                        upDocStatus.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        await upDocStatus.ExecuteNonQueryAsync();
                    }
                }

                if (actionLower == "hold" || actionLower == "reject")
                {
                    try
                    {
                        var authorId = await GetDocumentAuthorUserIdAsync(conn, dto.docId!);
                        if (!string.IsNullOrWhiteSpace(authorId)
                            && !string.Equals(authorId, approverId2, StringComparison.OrdinalIgnoreCase))
                        {
                            static string CombinePublicUrl(string baseUrl, string relativeUrl)
                            {
                                baseUrl = (baseUrl ?? string.Empty).Trim().TrimEnd('/');
                                relativeUrl = (relativeUrl ?? string.Empty).Trim();

                                if (string.IsNullOrWhiteSpace(baseUrl))
                                    baseUrl = "https://eboard.hyind.co.kr";

                                if (string.IsNullOrWhiteSpace(relativeUrl))
                                    relativeUrl = "/";

                                if (!relativeUrl.StartsWith("/", StringComparison.Ordinal))
                                    relativeUrl = "/" + relativeUrl;

                                return baseUrl + relativeUrl;
                            }

                            var publicBaseUrl =
                                (_cfg["App:PublicBaseUrl"]
                                 ?? _cfg["PublicBaseUrl"]
                                 ?? _cfg["WebPush:PublicBaseUrl"]
                                 ?? "https://eboard.hyind.co.kr").Trim();

                            var detailUrl = CombinePublicUrl(
                                publicBaseUrl,
                                "/Doc/DetailDX?id=" + Uri.EscapeDataString(dto.docId!));

                            var titleText = (_S?["PUSH_SummaryTitle"] ?? "PUSH_SummaryTitle").ToString();
                            var bodyText = actionLower == "hold"
                                ? (_S?["PUSH_ApprovalHold"] ?? "PUSH_ApprovalHold").ToString()
                                : (_S?["PUSH_ApprovalReject"] ?? "PUSH_ApprovalReject").ToString();

                            var tag = actionLower == "hold" ? "approval-author-hold" : "approval-author-reject";

                            await _webPushNotifier.SendToUserIdAsync(
                                userId: authorId.Trim(),
                                title: titleText,
                                body: bodyText,
                                url: detailUrl,
                                tag: tag);
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
            return await ApplyStampsAndReturnAsync(dto, newStatus, previewJson2);
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

            var currentCulture = DocControllerHelper.NormalizeCultureName(CultureInfo.CurrentUICulture.Name);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT d.DocId,
       d.TemplateCode,
       d.TemplateTitle,
       d.Status,
       d.DescriptorJson,
       d.OutputPath,
       d.CreatedAt,
       COALESCE(NULLIF(LTRIM(RTRIM(cm.TimeZoneId)), N''), N'Asia/Seoul') AS CompTimeZoneId
FROM dbo.Documents d
LEFT JOIN dbo.CompMasters cm ON cm.CompCd = d.CompCd
WHERE d.DocId = @id;";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

            string? descriptor = null, output = null, title = null, code = null, status = null, compTimeZoneId = null;
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
                    compTimeZoneId = rd["CompTimeZoneId"]?.ToString();
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
                        actedAtText = acted.HasValue ? FormatDetailLocalMinute(acted, compTimeZoneId, currentCulture) : null,
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
                        actedAtText = acted.HasValue ? FormatDetailLocalMinute(acted, compTimeZoneId, currentCulture) : null,
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
                createdAt = createdAt.HasValue ? FormatDetailLocalMinute(createdAt, compTimeZoneId, currentCulture) : null,
                descriptorJson = descriptor,
                previewJson = preview,
                approvals,
                cooperations
            });
        }

        private static string FormatDetailLocalMinute(DateTime? utcValue, string? compTimeZoneId, string? cultureName)
        {
            var local = DocControllerHelper.ConvertUtcToLocal(
                DocControllerHelper.TreatAsUtc(utcValue),
                compTimeZoneId
            );

            return DocControllerHelper.FormatLocalMinute(
                local,
                DocControllerHelper.NormalizeCultureName(cultureName)
            );
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

        // 2026.06.23 Removed: xlsx 재스캔 기반 렌더 진단 함수 제거 Contents DetailDX 렌더링 외 목표와 무관하여 제거한다

        // 2026.06.23 Added: DetailDX 작업 파일 복사 생성 Contents ComposeDX와 동일하게 원본을 직접 열지 않고 DocDXStamp 복사본을 DX Spreadsheet로 연다
        private string CreateDetailDxStampWorkbookCopy(string docId, string sourceAbsPath, string stampSessionId)
        {
            if (string.IsNullOrWhiteSpace(sourceAbsPath) || !System.IO.File.Exists(sourceAbsPath))
                return sourceAbsPath ?? string.Empty;

            var now = DateTime.Now;
            var safeDocId = SafeStampFilePart(docId);
            var safeSessionId = SafeStampFilePart(stampSessionId);
            var ext = Path.GetExtension(sourceAbsPath);

            if (string.IsNullOrWhiteSpace(ext))
                ext = ".xlsx";

            var dir = Path.Combine(
                _env.ContentRootPath,
                "App_Data",
                "DocDXStamp",
                now.ToString("yyyy", CultureInfo.InvariantCulture),
                now.ToString("MM", CultureInfo.InvariantCulture),
                now.ToString("dd", CultureInfo.InvariantCulture),
                safeDocId
            );

            Directory.CreateDirectory(dir);

            var destPath = Path.Combine(dir, safeDocId + "_" + safeSessionId + ext);

            lock (_detailDxStampFileLock)
            {
                System.IO.File.Copy(sourceAbsPath, destPath, overwrite: true);
            }

            return destPath;
        }

        [HttpPost("CleanupDetailDxStampTemp")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public IActionResult CleanupDetailDxStampTemp([FromBody] DetailDxStampTempCleanupRequest? request)
        {
            var docId = (request?.DocId ?? string.Empty).Trim();
            var stampWorkPath = (request?.StampWorkPath ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(stampWorkPath))
                return Json(new { ok = true, deleted = false });

            var deleted = TryDeleteDetailDxStampTempFile(docId, stampWorkPath);

            return Json(new { ok = true, deleted });
        }

        // 2026.06.23 Added: DetailDX embedded 누락 여부 확인 Contents NULL 또는 해시 불일치 대상이 있을 때만 클라이언트에서 1회 보정 요청한다
        private async Task<bool> HasPendingStampRowsAsync(SqlConnection conn, string docId, long? templateVersionId, string? descriptorJson)
        {
            var rows = await LoadStampRowsAsync(conn, docId, templateVersionId, descriptorJson);
            return rows.Any(x => x.ShouldEmbed);
        }

        // 2026.06.23 Added: 승인 협조 처리 후 스탬프 삽입 응답 생성 Contents Spreadsheet 세션 저장 성공 후에만 파일 교체 및 embedded 필드를 업데이트한다
        private async Task<IActionResult> ApplyStampsAndReturnAsync(ApproveDto dto, string status, string? previewJson = null)
        {
            var stampResult = await ApplyDocumentStampsFromSpreadsheetStateAsync(dto.docId, dto.spreadsheetState, dto.stampWorkPath);
            if (!string.IsNullOrWhiteSpace(previewJson))
                return Json(new { ok = true, docId = dto.docId, status, previewJson, stampApplied = stampResult.applied, stampSkipped = stampResult.skipped });

            return Json(new { ok = true, docId = dto.docId, status, stampApplied = stampResult.applied, stampSkipped = stampResult.skipped });
        }

        [HttpPost("BackfillStampsDX")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> BackfillStampsDX([FromBody] BackfillStampsDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto?.docId) || dto.spreadsheetState == null)
                return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" }, stage = "stamp-arg" });

            var result = await ApplyDocumentStampsFromSpreadsheetStateAsync(dto.docId, dto.spreadsheetState, dto.stampWorkPath);
            return Json(new { ok = result.ok, applied = result.applied, skipped = result.skipped, messages = result.messages });
        }

        // 2026.06.23 Changed: 현재 열린 DevExpress Spreadsheet 세션에 스탬프 이미지를 삽입 Contents 임시 PNG 파일 생성 없이 System.Drawing.Image 객체를 직접 AddPicture에 전달한다
        private async Task<(bool ok, int applied, int skipped, string[] messages)> ApplyDocumentStampsFromSpreadsheetStateAsync(string? docId, SpreadsheetClientState? spreadsheetState, string? stampWorkPath)
        {
            var safeDocId = (docId ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(safeDocId))
                return (false, 0, 0, new[] { "DOC_Err_BadRequest" });

            if (spreadsheetState == null)
                return (false, 0, 0, new[] { "DOC_Err_SaveFailed" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cs))
                return (false, 0, 0, new[] { "DOC_Err_RequestFailed" });

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            var docInfo = await LoadStampDocumentInfoAsync(conn, safeDocId);
            if (docInfo == null)
                return (false, 0, 0, new[] { "DOC_Err_DocumentNotFound" });

            var sourcePath = ResolveContentRootRelativePath(docInfo.OutputPath) ?? string.Empty;
            var workPath = ValidateStampWorkPath(stampWorkPath) ? stampWorkPath!.Trim() : string.Empty;
            if (string.IsNullOrWhiteSpace(workPath) || !System.IO.File.Exists(workPath))
                return (false, 0, 0, new[] { "DOC_Err_SaveFailed" });

            var targets = await LoadStampRowsAsync(conn, safeDocId, docInfo.TemplateVersionId, docInfo.DescriptorJson);
            targets = targets.Where(x => x.ShouldEmbed).ToList();
            if (targets.Count == 0)
                return (true, 0, 0, Array.Empty<string>());

            var generatedImages = new List<IDisposable>();

            try
            {
                var spreadsheet = SpreadsheetRequestProcessor.GetSpreadsheetFromState(spreadsheetState);
                if (spreadsheet == null)
                    return (false, 0, 0, new[] { "DOC_Err_SaveFailed" });

                var workbook = spreadsheet.Document;
                if (workbook == null || workbook.Worksheets.Count == 0)
                    return (false, 0, 0, new[] { "DOC_Err_SaveFailed" });

                var appliedTargets = new List<StampRow>();

                foreach (var target in targets)
                {
                    try
                    {
                        var (sheetName, localA1) = SplitSheetAndA1ForStamp(target.TargetA1);
                        if (string.IsNullOrWhiteSpace(localA1))
                        {
                            target.SkipReason = "target-empty";
                            continue;
                        }

                        DevExpress.Spreadsheet.Worksheet ws;
                        if (!string.IsNullOrWhiteSpace(sheetName))
                        {
                            DevExpress.Spreadsheet.Worksheet? matched = null;
                            for (var sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
                            {
                                var candidate = workbook.Worksheets[sheetIndex];
                                if (string.Equals(candidate.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                                {
                                    matched = candidate;
                                    break;
                                }
                            }

                            if (matched == null)
                            {
                                target.SkipReason = "worksheet-not-found";
                                continue;
                            }

                            ws = matched;
                        }
                        else
                        {
                            try { ws = workbook.Worksheets.ActiveWorksheet; }
                            catch { ws = workbook.Worksheets[0]; }
                        }

                        var range = ResolveStampTargetRange(ws, localA1);
                        var stampCanvasInfo = BuildStampCanvasInfo(ws, range);

                        // 2026.07.01 Changed: 기존 그림 삭제 범위 제한 Contents 같은 그림 이름만 삭제하여 다른 결재 협조 스탬프가 지워지지 않게 한다
                        var deletedPictureCount = DeleteExistingStampPictures(ws, target.PictureName);

                        // 2026.07.02 Changed: 날인 텍스트 분리 Contents 텍스트를 이미지에 합성하지 않고 엑셀 셀 값과 서식으로 표시한다
                        ApplyStampCaptionToRange(range, target.CaptionText, stampCanvasInfo);

                        using var stampImage = CreateTrimmedStampImageForInsert(target, stampCanvasInfo, out var stampLayoutInfo);
                        var stampImageForInsert = new GdiBitmap(stampImage);
                        generatedImages.Add(stampImageForInsert);

                        // 2026.07.02 Changed: 스탬프 삽입 방식 보정 Contents 잘라낸 이미지에 셀 비율 투명 여백을 추가한 뒤 결재란 범위에 삽입한다
                        var picture = ws.Pictures.AddPicture(stampImageForInsert, range);
                        picture.Name = target.PictureName;

                        _log.LogInformation(
                            "DetailDX stamp inserted docId={DocId} lineType={LineType} id={Id} a1={A1} sheet={SheetName} localA1={LocalA1} rangeTopRow={RangeTopRow} rangeBottomRow={RangeBottomRow} rangeLeftCol={RangeLeftCol} rangeRightCol={RangeRightCol} canvasWidth={CanvasWidth} canvasHeight={CanvasHeight} caption={Caption} sourcePath={SourcePath} sourceExists={SourceExists} sourceWidth={SourceWidth} sourceHeight={SourceHeight} trimLeft={TrimLeft} trimTop={TrimTop} trimRight={TrimRight} trimBottom={TrimBottom} trimWidth={TrimWidth} trimHeight={TrimHeight} deletedPictures={DeletedPictures} pictureName={PictureName}",
                            safeDocId,
                            target.LineType,
                            target.Id,
                            target.TargetA1,
                            ws.Name,
                            localA1,
                            range.TopRowIndex,
                            range.BottomRowIndex,
                            range.LeftColumnIndex,
                            range.RightColumnIndex,
                            stampCanvasInfo.WidthPx,
                            stampCanvasInfo.HeightPx,
                            target.CaptionText,
                            stampLayoutInfo.SourceImagePath,
                            stampLayoutInfo.SourceExists,
                            stampLayoutInfo.SourceWidthPx,
                            stampLayoutInfo.SourceHeightPx,
                            stampLayoutInfo.TrimLeftPx,
                            stampLayoutInfo.TrimTopPx,
                            stampLayoutInfo.TrimRightPx,
                            stampLayoutInfo.TrimBottomPx,
                            stampLayoutInfo.TrimWidthPx,
                            stampLayoutInfo.TrimHeightPx,
                            deletedPictureCount,
                            target.PictureName);

                        appliedTargets.Add(target);

                    }
                    catch (Exception exTarget)
                    {
                        _log.LogWarning(exTarget, "DetailDX stamp target skipped docId={DocId} lineType={LineType} id={Id} a1={A1}", safeDocId, target.LineType, target.Id, target.TargetA1);
                    }
                }

                if (appliedTargets.Count == 0)
                    return (true, 0, targets.Count, Array.Empty<string>());

                // 2026.06.23 Changed: byte accessor Open 문서는 SaveCopy 결과를 작업 파일에 기록 Contents 빈 세션이 원본을 덮어쓰지 않도록 applied 대상이 있을 때만 저장한다
                var savedBytes = spreadsheet.SaveCopy(DevExpress.Spreadsheet.DocumentFormat.Xlsx);
                if (savedBytes == null || savedBytes.Length <= 0)
                    return (false, appliedTargets.Count, targets.Count - appliedTargets.Count, new[] { "DOC_Err_SaveFailed" });

                System.IO.File.WriteAllBytes(workPath, savedBytes);

                ReplaceSourceDocumentFromStampWorkFile(sourcePath, workPath, safeDocId);
                await UpdateStampEmbeddedRowsAsync(conn, appliedTargets);

                return (true, appliedTargets.Count, targets.Count - appliedTargets.Count, Array.Empty<string>());
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "DetailDX stamp apply failed docId={DocId}", safeDocId);
                return (false, 0, targets.Count, new[] { "DOC_Err_SaveFailed" });
            }
            finally
            {
                foreach (var image in generatedImages)
                {
                    try { image.Dispose(); } catch { }
                }
            }
        }

        private async Task<StampDocumentInfo?> LoadStampDocumentInfoAsync(SqlConnection conn, string docId)
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT TOP 1 DocId, OutputPath, DescriptorJson, TemplateVersionId
FROM dbo.Documents
WHERE DocId = @DocId;";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

            await using var rd = await cmd.ExecuteReaderAsync();
            if (!await rd.ReadAsync())
                return null;

            return new StampDocumentInfo
            {
                DocId = rd["DocId"]?.ToString() ?? string.Empty,
                OutputPath = rd["OutputPath"]?.ToString() ?? string.Empty,
                DescriptorJson = rd["DescriptorJson"]?.ToString() ?? "{}",
                TemplateVersionId = rd["TemplateVersionId"] == DBNull.Value ? null : Convert.ToInt64(rd["TemplateVersionId"], CultureInfo.InvariantCulture)
            };
        }

        private async Task<List<StampRow>> LoadStampRowsAsync(SqlConnection conn, string docId, long? templateVersionId, string? descriptorJson)
        {
            var rows = new List<StampRow>();

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT a.Id,
       a.StepOrder,
       a.RoleKey,
       a.UserId,
       a.ApproverValue,
       a.Status,
       a.Action,
       a.ActedAt,
       a.ApproverDisplayText,
       a.SignaturePath,
       a.StampEmbeddedAt,
       a.StampStateHash,
       ISNULL(ta.CellA1, ta.A1) AS TargetA1
FROM dbo.DocumentApprovals a
LEFT JOIN dbo.DocTemplateApproval ta
       ON ta.VersionId = @VersionId
      AND ta.Slot = a.StepOrder
WHERE a.DocId = @DocId
ORDER BY a.StepOrder;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = (object?)templateVersionId ?? DBNull.Value });

                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    rows.Add(new StampRow
                    {
                        LineType = ApprovalLineType,
                        Id = Convert.ToInt64(rd["Id"], CultureInfo.InvariantCulture),
                        StepOrder = rd["StepOrder"] == DBNull.Value ? null : Convert.ToInt32(rd["StepOrder"], CultureInfo.InvariantCulture),
                        RoleKey = rd["RoleKey"]?.ToString() ?? string.Empty,
                        UserId = rd["UserId"]?.ToString() ?? string.Empty,
                        ApproverValue = rd["ApproverValue"]?.ToString() ?? string.Empty,
                        Status = rd["Status"]?.ToString() ?? string.Empty,
                        Action = rd["Action"]?.ToString() ?? string.Empty,
                        ActedAtUtc = rd["ActedAt"] == DBNull.Value ? null : Convert.ToDateTime(rd["ActedAt"], CultureInfo.InvariantCulture),
                        DisplayText = rd["ApproverDisplayText"]?.ToString() ?? string.Empty,
                        SignaturePath = rd["SignaturePath"]?.ToString() ?? string.Empty,
                        StampEmbeddedAt = rd["StampEmbeddedAt"] == DBNull.Value ? null : Convert.ToDateTime(rd["StampEmbeddedAt"], CultureInfo.InvariantCulture),
                        StampStateHash = rd["StampStateHash"]?.ToString() ?? string.Empty,
                        TargetA1 = rd["TargetA1"]?.ToString() ?? string.Empty
                    });
                }
            }

            var templateDescriptorJson = "{}";
            if (templateVersionId.HasValue)
            {
                try
                {
                    await using var td = conn.CreateCommand();
                    td.CommandText = "SELECT TOP 1 DescriptorJson FROM dbo.DocTemplateVersion WHERE Id = @Id;";
                    td.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = templateVersionId.Value });
                    var obj = await td.ExecuteScalarAsync();
                    templateDescriptorJson = obj == null || obj == DBNull.Value ? "{}" : obj.ToString() ?? "{}";
                }
                catch
                {
                    templateDescriptorJson = "{}";
                }
            }

            var coopMaps = BuildStampCooperationCellMaps(descriptorJson, templateDescriptorJson);

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT dc.Id,
       dc.RoleKey,
       dc.UserId,
       dc.ApproverValue,
       dc.Status,
       dc.Action,
       dc.ActedAt,
       dc.ApproverDisplayText,
       COALESCE(NULLIF(LTRIM(RTRIM(ISNULL(dc.SignaturePath, N''))), N''), up.SignatureRelativePath, N'') AS SignaturePath,
       dc.StampEmbeddedAt,
       dc.StampStateHash
FROM dbo.DocumentCooperations dc
LEFT JOIN dbo.UserProfiles up ON up.UserId = dc.UserId
WHERE dc.DocId = @DocId
ORDER BY dc.RoleKey;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    var roleKey = rd["RoleKey"]?.ToString() ?? string.Empty;
                    var userId = rd["UserId"]?.ToString() ?? string.Empty;
                    var approverValue = rd["ApproverValue"]?.ToString() ?? string.Empty;

                    var targetA1 = string.Empty;
                    if (!string.IsNullOrWhiteSpace(roleKey) && coopMaps.ByRoleKey.TryGetValue(roleKey, out var byRole))
                        targetA1 = byRole;
                    else if (!string.IsNullOrWhiteSpace(userId) && coopMaps.ByApprover.TryGetValue(userId, out var byUser))
                        targetA1 = byUser;
                    else if (!string.IsNullOrWhiteSpace(approverValue) && coopMaps.ByApprover.TryGetValue(approverValue, out var byValue))
                        targetA1 = byValue;

                    rows.Add(new StampRow
                    {
                        LineType = CooperationLineType,
                        Id = Convert.ToInt64(rd["Id"], CultureInfo.InvariantCulture),
                        StepOrder = ParseRoleOrder(roleKey),
                        RoleKey = roleKey,
                        UserId = userId,
                        ApproverValue = approverValue,
                        Status = rd["Status"]?.ToString() ?? string.Empty,
                        Action = rd["Action"]?.ToString() ?? string.Empty,
                        ActedAtUtc = rd["ActedAt"] == DBNull.Value ? null : Convert.ToDateTime(rd["ActedAt"], CultureInfo.InvariantCulture),
                        DisplayText = rd["ApproverDisplayText"]?.ToString() ?? string.Empty,
                        SignaturePath = rd["SignaturePath"]?.ToString() ?? string.Empty,
                        StampEmbeddedAt = rd["StampEmbeddedAt"] == DBNull.Value ? null : Convert.ToDateTime(rd["StampEmbeddedAt"], CultureInfo.InvariantCulture),
                        StampStateHash = rd["StampStateHash"]?.ToString() ?? string.Empty,
                        TargetA1 = targetA1
                    });
                }
            }

            foreach (var row in rows)
            {
                ResolveStampRow(row);
                _log.LogInformation(
                    "DetailDX stamp row resolved docId={DocId} lineType={LineType} id={Id} roleKey={RoleKey} step={StepOrder} status={Status} action={Action} targetA1={TargetA1} displayText={DisplayText} caption={Caption} signaturePath={SignaturePath} sourcePath={SourcePath} sourceExists={SourceExists} shouldEmbed={ShouldEmbed} skipReason={SkipReason} oldHash={OldHash} newHash={NewHash}",
                    docId,
                    row.LineType,
                    row.Id,
                    row.RoleKey,
                    row.StepOrder,
                    row.Status,
                    row.Action,
                    row.TargetA1,
                    row.DisplayText,
                    row.CaptionText,
                    row.SignaturePath,
                    row.SourceImagePath,
                    !string.IsNullOrWhiteSpace(row.SourceImagePath) && System.IO.File.Exists(row.SourceImagePath),
                    row.ShouldEmbed,
                    row.SkipReason,
                    row.StampStateHash,
                    row.NewStateHash);
            }

            return rows;
        }

        private void ResolveStampRow(StampRow row)
        {
            var kind = ResolveStampKind(row.Status, row.Action);
            if (kind == null)
            {
                row.SkipReason = "not-stamp-status";
                return;
            }

            if (string.IsNullOrWhiteSpace(row.TargetA1))
            {
                row.SkipReason = "target-empty";
                return;
            }

            row.StampKind = kind.Value;
            row.CaptionText = FormatStampCaption(row.DisplayText);

            var imagePath = ResolveCommonStampPath(kind.Value);
            if (kind.Value == StampKind.Approved)
            {
                var signaturePath = ResolveSignaturePhysicalPath(row.SignaturePath);
                if (!string.IsNullOrWhiteSpace(signaturePath) && System.IO.File.Exists(signaturePath))
                    imagePath = signaturePath;
            }

            if (string.IsNullOrWhiteSpace(imagePath) || !System.IO.File.Exists(imagePath))
            {
                row.SkipReason = "stamp-image-not-found";
                return;
            }

            row.SourceImagePath = imagePath;
            row.PictureName = BuildStampPictureName(row);
            row.NewStateHash = BuildStampStateHash(row);
            row.ShouldEmbed = row.StampEmbeddedAt == null || !string.Equals(row.StampStateHash ?? string.Empty, row.NewStateHash, StringComparison.OrdinalIgnoreCase);
        }

        private static StampKind? ResolveStampKind(string? status, string? action)
        {
            var s = (status ?? string.Empty).Trim().ToUpperInvariant();
            var a = (action ?? string.Empty).Trim().ToUpperInvariant();

            if (a == "RECALLED" || s == "RECALLED") return null;
            if (a == "HOLD" || s == "HOLD" || s == "ONHOLD" || s.StartsWith("PENDINGHOLD", StringComparison.OrdinalIgnoreCase)) return StampKind.Holding;
            if (a == "REJECT" || a == "COOPERATEREJECT" || s == "REJECTED" || s.StartsWith("REJECTED", StringComparison.OrdinalIgnoreCase)) return StampKind.Rejected;
            if (a == "APPROVE" || a == "COOPERATE" || s == "APPROVED" || s == "COOPERATED") return StampKind.Approved;
            return null;
        }

        // 2026.06.23 Added: 스탬프 대상 범위를 병합셀 전체 범위로 보정 Contents 단일 셀 주소가 병합 영역 안에 있을 때 이미지가 1칸 크기로 삽입되는 문제를 방지한다
        private static CellRange ResolveStampTargetRange(DevExpress.Spreadsheet.Worksheet worksheet, string localA1)
        {
            var range = worksheet.Range[localA1];
            var mergedRanges = range.GetMergedRanges();
            if (mergedRanges != null && mergedRanges.Count > 0)
                return mergedRanges[0];

            return range;
        }

        // 2026.07.01 Changed: 기존 스탬프 그림 삭제 기준 수정 Contents 현재 대상과 같은 그림 이름만 삭제하고 겹침 범위 삭제는 사용하지 않는다
        private static int DeleteExistingStampPictures(DevExpress.Spreadsheet.Worksheet worksheet, string pictureName)
        {
            if (string.IsNullOrWhiteSpace(pictureName))
                return 0;

            var deleted = 0;

            try
            {
                var oldPictures = worksheet.Pictures.GetPicturesByName(pictureName);
                for (var picIndex = oldPictures.Count - 1; picIndex >= 0; picIndex--)
                {
                    try
                    {
                        oldPictures[picIndex].Delete();
                        deleted++;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return deleted;
        }


        // 2026.07.01 Changed: 스탬프 대상 승인셀 크기 계산 Contents DevExpress 행 높이 단위가 quarter point 계열로 들어오는 경우를 픽셀로 정상 환산한다
        private static StampCanvasInfo BuildStampCanvasInfo(DevExpress.Spreadsheet.Worksheet worksheet, CellRange range)
        {
            var widthPx = 0d;
            var heightPx = 0d;

            for (var colIndex = range.LeftColumnIndex; colIndex <= range.RightColumnIndex; colIndex++)
            {
                try
                {
                    var column = worksheet.Columns[colIndex];
                    if (!column.Visible)
                        continue;

                    var columnWidth = Convert.ToDouble(column.WidthInPixels, CultureInfo.InvariantCulture);
                    if (columnWidth > 0)
                        widthPx += columnWidth;
                }
                catch
                {
                }
            }

            for (var rowIndex = range.TopRowIndex; rowIndex <= range.BottomRowIndex; rowIndex++)
            {
                try
                {
                    var row = worksheet.Rows[rowIndex];
                    if (!row.Visible)
                        continue;

                    var rowHeightPx = ResolveSpreadsheetRowHeightPx(row);
                    if (rowHeightPx > 0)
                        heightPx += rowHeightPx;
                }
                catch
                {
                }
            }

            return new StampCanvasInfo
            {
                WidthPx = ClampStampCanvasSize(widthPx, 24, 1600, 96),
                HeightPx = ClampStampCanvasSize(heightPx, 24, 1600, 80)
            };
        }

        // 2026.07.01 Added: DevExpress 행 높이 픽셀 환산 Contents row Height 값이 quarter point 계열이면 4로 나눈 뒤 96 DPI 기준 픽셀로 변환한다
        private static double ResolveSpreadsheetRowHeightPx(DevExpress.Spreadsheet.Row row)
        {
            if (TryGetNumericPropertyValue(row, "HeightInPixels", out var heightInPixels) && heightInPixels > 0d)
                return heightInPixels;

            if (TryGetNumericPropertyValue(row, "HeightInPoints", out var heightInPoints) && heightInPoints > 0d)
                return heightInPoints * 96d / 72d;

            var rawHeight = Convert.ToDouble(row.Height, CultureInfo.InvariantCulture);
            if (double.IsNaN(rawHeight) || double.IsInfinity(rawHeight) || rawHeight <= 0d)
                return 0d;

            if (rawHeight > 40d)
                return rawHeight / 4d * 96d / 72d;

            return rawHeight * 96d / 72d;
        }

        // 2026.07.01 Added: 숫자 속성 조회 Contents DevExpress 버전별 행 높이 속성 차이를 안전하게 처리한다
        private static bool TryGetNumericPropertyValue(object source, string propertyName, out double value)
        {
            value = 0d;

            if (source == null || string.IsNullOrWhiteSpace(propertyName))
                return false;

            try
            {
                var property = source.GetType().GetProperty(propertyName);
                if (property == null)
                    return false;

                var raw = property.GetValue(source);
                if (raw == null)
                    return false;

                value = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return !double.IsNaN(value) && !double.IsInfinity(value);
            }
            catch
            {
                value = 0d;
                return false;
            }
        }

        private static int ClampStampCanvasSize(double value, int min, int max, int fallback)
        {
            if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0)
                return fallback;

            var rounded = (int)Math.Ceiling(value);
            if (rounded < min) return min;
            if (rounded > max) return max;
            return rounded;
        }
        // 2026.07.02 Changed: 스탬프 이미지 생성 보정 Contents 이진화로 찾은 실제 스탬프 경계를 기준으로 원본 비율을 유지하고 텍스트 영역 예약 없이 셀 이미지 영역 전체에 최대 배치한다
        // 2026.07.02 Changed: 스탬프 이미지 처리 보정 Contents 원본 이미지에서 실제 내용 영역만 잘라낸 뒤 셀 비율 투명 여백을 추가한다
        private GdiBitmap CreateTrimmedStampImageForInsert(StampRow row, StampCanvasInfo canvasInfo, out StampImageLayoutInfo layoutInfo)
        {
            using var sourceImage = GdiImage.FromFile(row.SourceImagePath);
            using var source = new GdiBitmap(sourceImage);

            var trimBounds = FindStampContentBounds(source);
            if (trimBounds.Width <= 0 || trimBounds.Height <= 0)
                trimBounds = new GdiRectangle(0, 0, source.Width, source.Height);

            layoutInfo = new StampImageLayoutInfo
            {
                SourceImagePath = row.SourceImagePath,
                SourceExists = !string.IsNullOrWhiteSpace(row.SourceImagePath) && System.IO.File.Exists(row.SourceImagePath),
                SourceWidthPx = source.Width,
                SourceHeightPx = source.Height,
                TrimLeftPx = trimBounds.Left,
                TrimTopPx = trimBounds.Top,
                TrimRightPx = trimBounds.Right - 1,
                TrimBottomPx = trimBounds.Bottom - 1,
                TrimWidthPx = trimBounds.Width,
                TrimHeightPx = trimBounds.Height
            };

            using var trimmed = source.Clone(trimBounds, GdiPixelFormat.Format32bppArgb);
            return CreateStampImageWithCellPadding(trimmed, canvasInfo, 0.9d);
        }

        // 2026.07.02 Changed: 스탬프 외곽 여백 추가 보정 Contents 잘라낸 원본 픽셀은 변경하지 않고 셀 비율의 투명 여백을 추가하여 셀 영역의 90퍼센트 안에 표시한다
        private static GdiBitmap CreateStampImageWithCellPadding(GdiBitmap source, StampCanvasInfo canvasInfo, double contentRatio)
        {
            if (source.Width <= 0 || source.Height <= 0)
                return new GdiBitmap(1, 1, GdiPixelFormat.Format32bppArgb);

            if (double.IsNaN(contentRatio) || double.IsInfinity(contentRatio) || contentRatio <= 0d || contentRatio > 1d)
                contentRatio = 1d;

            var targetAspect = Math.Max(1, canvasInfo.WidthPx) / (double)Math.Max(1, canvasInfo.HeightPx);
            var sourceAspect = source.Width / (double)source.Height;

            var canvasWidth = sourceAspect >= targetAspect
                ? (int)Math.Ceiling(source.Width / contentRatio)
                : (int)Math.Ceiling((source.Height / contentRatio) * targetAspect);

            var canvasHeight = sourceAspect >= targetAspect
                ? (int)Math.Ceiling(canvasWidth / targetAspect)
                : (int)Math.Ceiling(source.Height / contentRatio);

            canvasWidth = Math.Max(source.Width, canvasWidth);
            canvasHeight = Math.Max(source.Height, canvasHeight);

            var offsetX = Math.Max(0, (canvasWidth - source.Width) / 2);
            var offsetY = Math.Max(0, (canvasHeight - source.Height) / 2);

            var canvas = new GdiBitmap(canvasWidth, canvasHeight, GdiPixelFormat.Format32bppArgb);

            for (var y = 0; y < canvasHeight; y++)
                for (var x = 0; x < canvasWidth; x++)
                    canvas.SetPixel(x, y, GdiColor.Transparent);

            // 2026.07.02 Added: 원본 픽셀 복사 Contents 잘라낸 이미지의 픽셀 값을 그대로 복사하고 크기 색상 투명도는 변경하지 않는다
            for (var y = 0; y < source.Height; y++)
                for (var x = 0; x < source.Width; x++)
                    canvas.SetPixel(offsetX + x, offsetY + y, source.GetPixel(x, y));

            return canvas;
        }

        // 2026.07.02 Added: 스탬프 실제 내용 영역 계산 Contents 바깥 여백만 잘라내기 위해 이미지 가장자리 배경과 다른 첫 픽셀 위치를 찾는다
        private static GdiRectangle FindStampContentBounds(GdiBitmap source)
        {
            if (source.Width <= 0 || source.Height <= 0)
                return GdiRectangle.Empty;

            if (HasTransparentMargin(source))
                return FindOpaqueBounds(source);

            var backgroundColor = GetCornerAverageColor(source);

            var top = 0;
            while (top < source.Height && IsBackgroundRow(source, top, backgroundColor))
                top++;

            var bottom = source.Height - 1;
            while (bottom >= top && IsBackgroundRow(source, bottom, backgroundColor))
                bottom--;

            var left = 0;
            while (left < source.Width && IsBackgroundColumn(source, left, top, bottom, backgroundColor))
                left++;

            var right = source.Width - 1;
            while (right >= left && IsBackgroundColumn(source, right, top, bottom, backgroundColor))
                right--;

            if (left > right || top > bottom)
                return new GdiRectangle(0, 0, source.Width, source.Height);

            return GdiRectangle.FromLTRB(left, top, right + 1, bottom + 1);
        }

        // 2026.07.02 Added: 투명 여백 우선 판단 Contents 투명 배경 서명 파일은 알파값 기준으로 바깥 여백만 잘라낸다
        private static bool HasTransparentMargin(GdiBitmap source)
        {
            for (var x = 0; x < source.Width; x++)
            {
                if (source.GetPixel(x, 0).A == 0 || source.GetPixel(x, source.Height - 1).A == 0)
                    return true;
            }

            for (var y = 0; y < source.Height; y++)
            {
                if (source.GetPixel(0, y).A == 0 || source.GetPixel(source.Width - 1, y).A == 0)
                    return true;
            }

            return false;
        }

        // 2026.07.02 Added: 투명 배경 내용 영역 계산 Contents 알파값이 있는 픽셀만 실제 내용으로 간주한다
        private static GdiRectangle FindOpaqueBounds(GdiBitmap source)
        {
            var left = source.Width;
            var top = source.Height;
            var right = -1;
            var bottom = -1;

            for (var y = 0; y < source.Height; y++)
            {
                for (var x = 0; x < source.Width; x++)
                {
                    if (source.GetPixel(x, y).A == 0)
                        continue;

                    if (x < left) left = x;
                    if (y < top) top = y;
                    if (x > right) right = x;
                    if (y > bottom) bottom = y;
                }
            }

            if (right < left || bottom < top)
                return new GdiRectangle(0, 0, source.Width, source.Height);

            return GdiRectangle.FromLTRB(left, top, right + 1, bottom + 1);
        }

        // 2026.07.02 Added: 모서리 평균 배경색 계산 Contents 흰 종이 스캔 이미지의 바깥 여백을 제거하기 위한 기준색을 만든다
        private static GdiColor GetCornerAverageColor(GdiBitmap source)
        {
            var samples = new[]
            {
                source.GetPixel(0, 0),
                source.GetPixel(source.Width - 1, 0),
                source.GetPixel(0, source.Height - 1),
                source.GetPixel(source.Width - 1, source.Height - 1)
            };

            var a = (int)Math.Round(samples.Average(x => x.A));
            var r = (int)Math.Round(samples.Average(x => x.R));
            var g = (int)Math.Round(samples.Average(x => x.G));
            var b = (int)Math.Round(samples.Average(x => x.B));

            return GdiColor.FromArgb(a, r, g, b);
        }

        // 2026.07.02 Added: 배경 행 판단 Contents 행 전체가 모서리 배경색과 같으면 바깥 여백으로 본다
        private static bool IsBackgroundRow(GdiBitmap source, int rowIndex, GdiColor backgroundColor)
        {
            for (var x = 0; x < source.Width; x++)
            {
                if (!IsBackgroundLikePixel(source.GetPixel(x, rowIndex), backgroundColor))
                    return false;
            }

            return true;
        }

        // 2026.07.02 Added: 배경 열 판단 Contents 열 전체가 모서리 배경색과 같으면 바깥 여백으로 본다
        private static bool IsBackgroundColumn(GdiBitmap source, int colIndex, int top, int bottom, GdiColor backgroundColor)
        {
            for (var y = top; y <= bottom; y++)
            {
                if (!IsBackgroundLikePixel(source.GetPixel(colIndex, y), backgroundColor))
                    return false;
            }

            return true;
        }

        // 2026.07.02 Added: 배경 유사 픽셀 판단 Contents 색상 비교는 여백 판정에만 사용하고 실제 출력 픽셀은 변경하지 않는다
        private static bool IsBackgroundLikePixel(GdiColor color, GdiColor backgroundColor)
        {
            if (color.A == 0)
                return true;

            return Math.Abs(color.A - backgroundColor.A) <= 24
                && Math.Abs(color.R - backgroundColor.R) <= 24
                && Math.Abs(color.G - backgroundColor.G) <= 24
                && Math.Abs(color.B - backgroundColor.B) <= 24;
        }

        // 2026.07.01 Changed: 날인 텍스트 셀 글꼴 크기 계산 Contents 엑셀 기본 11포인트를 우선 사용하고 셀이 작을 때만 축소한다
        private static double ResolveStampCaptionFontSizePt(StampCanvasInfo canvasInfo)
        {
            const double excelDefaultFontSizePt = 11d;
            const double minFontSizePt = 8d;

            var cellHeightPt = Math.Max(1d, canvasInfo.HeightPx * 72d / 96d);
            if (cellHeightPt >= 32d)
                return excelDefaultFontSizePt;

            var maxByHeight = Math.Max(minFontSizePt, cellHeightPt * 0.32d);
            return Math.Max(minFontSizePt, Math.Min(excelDefaultFontSizePt, maxByHeight));
        }

        // 2026.07.01 Added: 날인 텍스트 셀 표시 Contents 텍스트를 bitmap에 합성하지 않고 엑셀 셀 값과 서식으로 표시한다
        private static void ApplyStampCaptionToRange(CellRange range, string? caption, StampCanvasInfo canvasInfo)
        {
            var text = (caption ?? string.Empty).Trim();

            try
            {
                range.Value = text;
                range.Font.Name = "Malgun Gothic";
                range.Font.Size = ResolveStampCaptionFontSizePt(canvasInfo);
                range.Font.Bold = false;
                range.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                range.Alignment.Vertical = SpreadsheetVerticalAlignment.Bottom;
                range.Alignment.WrapText = false;
            }
            catch
            {
            }
        }



        private void ReplaceSourceDocumentFromStampWorkFile(string sourcePath, string workPath, string docId)
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || string.IsNullOrWhiteSpace(workPath))
                return;

            if (!System.IO.File.Exists(workPath))
                return;

            lock (_detailDxStampFileLock)
            {
                System.IO.File.Copy(workPath, sourcePath, overwrite: true);
            }
        }

        private async Task UpdateStampEmbeddedRowsAsync(SqlConnection conn, IReadOnlyList<StampRow> rows)
        {
            if (rows.Count == 0) return;
            await using var tx = await conn.BeginTransactionAsync();
            try
            {
                foreach (var row in rows)
                {
                    if (string.Equals(row.LineType, ApprovalLineType, StringComparison.OrdinalIgnoreCase))
                    {
                        await using var cmd = conn.CreateCommand();
                        cmd.Transaction = (SqlTransaction)tx;
                        cmd.CommandText = @"
UPDATE dbo.DocumentApprovals
SET StampEmbeddedAt = SYSUTCDATETIME(),
    StampStateHash = @Hash
WHERE Id = @Id;";
                        cmd.Parameters.Add(new SqlParameter("@Hash", SqlDbType.VarChar, 64) { Value = row.NewStateHash });
                        cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = row.Id });
                        await cmd.ExecuteNonQueryAsync();
                    }
                    else if (string.Equals(row.LineType, CooperationLineType, StringComparison.OrdinalIgnoreCase))
                    {
                        await using var cmd = conn.CreateCommand();
                        cmd.Transaction = (SqlTransaction)tx;
                        cmd.CommandText = @"
UPDATE dbo.DocumentCooperations
SET StampEmbeddedAt = SYSUTCDATETIME(),
    StampStateHash = @Hash,
    SignaturePath = CASE
                        WHEN NULLIF(LTRIM(RTRIM(ISNULL(SignaturePath, N''))), N'') IS NULL
                             AND NULLIF(LTRIM(RTRIM(ISNULL(@SignaturePath, N''))), N'') IS NOT NULL
                        THEN @SignaturePath
                        ELSE SignaturePath
                    END
WHERE Id = @Id;";
                        cmd.Parameters.Add(new SqlParameter("@Hash", SqlDbType.VarChar, 64) { Value = row.NewStateHash });
                        cmd.Parameters.Add(new SqlParameter("@SignaturePath", SqlDbType.NVarChar, 400) { Value = (object?)row.SignaturePath ?? DBNull.Value });
                        cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = row.Id });
                        await cmd.ExecuteNonQueryAsync();
                    }
                }

                await tx.CommitAsync();
            }
            catch
            {
                try { await tx.RollbackAsync(); } catch { }
                throw;
            }
        }

        private string ResolveCommonStampPath(StampKind kind)
        {
            var fileName = kind switch
            {
                StampKind.Approved => "Approved.png",
                StampKind.Holding => "Holding.png",
                StampKind.Rejected => "Rejected.png",
                _ => "Approved.png"
            };
            return Path.Combine(_env.ContentRootPath, "App_Data", "Signatures", fileName);
        }

        private string ResolveSignaturePhysicalPath(string? relativePath)
        {
            var raw = (relativePath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            if (Path.IsPathRooted(raw)) return raw;

            var normalized = raw.TrimStart('\\', '/').Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
            if (normalized.StartsWith("App_Data" + Path.DirectorySeparatorChar + "Signatures", StringComparison.OrdinalIgnoreCase))
                return Path.Combine(_env.ContentRootPath, normalized);

            return Path.Combine(_env.ContentRootPath, "App_Data", "Signatures", normalized);
        }

        private static string BuildStampPictureName(StampRow row)
        {
            var line = string.Equals(row.LineType, CooperationLineType, StringComparison.OrdinalIgnoreCase) ? "COOPERATION" : "APPROVAL";
            return StampPicturePrefix + line + "_" + row.Id.ToString(CultureInfo.InvariantCulture);
        }

        private static string BuildStampStateHash(StampRow row)
        {
            var imageTicks = 0L;
            try
            {
                if (!string.IsNullOrWhiteSpace(row.SourceImagePath) && System.IO.File.Exists(row.SourceImagePath))
                    imageTicks = System.IO.File.GetLastWriteTimeUtc(row.SourceImagePath).Ticks;
            }
            catch
            {
                imageTicks = 0L;
            }

            var raw = string.Join("\u001F", new[]
            {
                StampRuleVersion,
                row.LineType,
                row.Id.ToString(CultureInfo.InvariantCulture),
                row.StepOrder?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                row.RoleKey,
                row.Status,
                row.Action,
                row.ActedAtUtc?.ToString("O", CultureInfo.InvariantCulture) ?? string.Empty,
                row.DisplayText,
                row.SignaturePath,
                row.SourceImagePath,
                imageTicks.ToString(CultureInfo.InvariantCulture),
                row.TargetA1,
                row.StampKind?.ToString() ?? string.Empty
            });
            using var sha = System.Security.Cryptography.SHA256.Create();
            return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(raw))).ToLowerInvariant();
        }

        private static string FormatStampCaption(string? displayText)
        {
            var parts = Regex.Split((displayText ?? string.Empty).Trim(), @"\s+").Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            if (parts.Length >= 3) return string.Join(" ", parts.Skip(parts.Length - 2));
            if (parts.Length == 2) return string.Join(" ", parts);
            return parts.Length == 1 ? parts[0] : string.Empty;
        }

        private static (string sheetName, string localA1) SplitSheetAndA1ForStamp(string? a1)
        {
            var s = (a1 ?? string.Empty).Trim().Replace("$", string.Empty).Replace("'", string.Empty);
            if (string.IsNullOrWhiteSpace(s)) return (string.Empty, string.Empty);
            var bang = s.LastIndexOf('!');
            return bang >= 0 ? (s[..bang].Trim(), s[(bang + 1)..].Trim()) : (string.Empty, s);
        }

        private static CooperationStampCellMaps BuildStampCooperationCellMaps(string? documentDescriptorJson, string? templateDescriptorJson)
        {
            var maps = new CooperationStampCellMaps();
            AppendStampCooperationCells(documentDescriptorJson, maps);
            if (maps.ByRoleKey.Count == 0) AppendStampCooperationCells(templateDescriptorJson, maps);
            return maps;
        }

        private static void AppendStampCooperationCells(string? descriptorJson, CooperationStampCellMaps maps)
        {
            if (string.IsNullOrWhiteSpace(descriptorJson) || descriptorJson.Trim() == "{}") return;
            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                if (!(doc.RootElement.TryGetProperty("Cooperations", out var coops) || doc.RootElement.TryGetProperty("cooperations", out coops)) || coops.ValueKind != JsonValueKind.Array)
                    return;

                var seq = 1;
                foreach (var coop in coops.EnumerateArray())
                {
                    var roleKey = TryGetJsonString(coop, "RoleKey") ?? TryGetJsonString(coop, "roleKey") ?? "C" + seq.ToString(CultureInfo.InvariantCulture);
                    var approver = TryGetJsonString(coop, "ApproverValue") ?? TryGetJsonString(coop, "approverValue") ?? TryGetJsonString(coop, "value") ?? string.Empty;
                    var a1 = TryGetStampCooperationCellA1(coop);
                    if (!string.IsNullOrWhiteSpace(roleKey) && !string.IsNullOrWhiteSpace(a1)) maps.ByRoleKey[roleKey] = a1;
                    if (!string.IsNullOrWhiteSpace(approver) && !string.IsNullOrWhiteSpace(a1)) maps.ByApprover[approver] = a1;
                    seq++;
                }
            }
            catch
            {
            }
        }

        private static string TryGetStampCooperationCellA1(JsonElement coop)
        {
            if (coop.TryGetProperty("Cell", out var cell) && cell.ValueKind == JsonValueKind.Object)
            {
                var a1 = TryGetJsonString(cell, "A1") ?? TryGetJsonString(cell, "a1") ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(a1)) return a1;
            }
            if (coop.TryGetProperty("cell", out var cell2) && cell2.ValueKind == JsonValueKind.Object)
            {
                var a1 = TryGetJsonString(cell2, "A1") ?? TryGetJsonString(cell2, "a1") ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(a1)) return a1;
            }
            return TryGetJsonString(coop, "A1") ?? TryGetJsonString(coop, "a1") ?? TryGetJsonString(coop, "cellA1") ?? TryGetJsonString(coop, "CellA1") ?? string.Empty;
        }

        private bool ValidateStampWorkPath(string? path)
        {
            var raw = (path ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw) || !Path.IsPathRooted(raw)) return false;
            var root = Path.GetFullPath(Path.Combine(_env.ContentRootPath, "App_Data", "DocDXStamp"));
            var full = Path.GetFullPath(raw);
            return full.StartsWith(root, StringComparison.OrdinalIgnoreCase);
        }

        private static string SafeStampFilePart(string? value)
        {
            var s = string.IsNullOrWhiteSpace(value) ? "none" : value.Trim();
            foreach (var ch in Path.GetInvalidFileNameChars()) s = s.Replace(ch, '_');
            return Regex.Replace(s, @"[^a-zA-Z0-9_\-.]", "_");
        }

        private static void TryDeleteStampFile(string? path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && System.IO.File.Exists(path))
                    System.IO.File.Delete(path);
            }
            catch
            {
            }
        }

        private sealed class StampDocumentInfo
        {
            public string DocId { get; set; } = string.Empty;
            public string OutputPath { get; set; } = string.Empty;
            public string DescriptorJson { get; set; } = "{}";
            public long? TemplateVersionId { get; set; }
        }

        private sealed class StampRow
        {
            public string LineType { get; set; } = string.Empty;
            public long Id { get; set; }
            public int? StepOrder { get; set; }
            public string RoleKey { get; set; } = string.Empty;
            public string UserId { get; set; } = string.Empty;
            public string ApproverValue { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string Action { get; set; } = string.Empty;
            public DateTime? ActedAtUtc { get; set; }
            public string DisplayText { get; set; } = string.Empty;
            public string CaptionText { get; set; } = string.Empty;
            public string SignaturePath { get; set; } = string.Empty;
            public DateTime? StampEmbeddedAt { get; set; }
            public string StampStateHash { get; set; } = string.Empty;
            public string NewStateHash { get; set; } = string.Empty;
            public string TargetA1 { get; set; } = string.Empty;
            public string SourceImagePath { get; set; } = string.Empty;
            public StampKind? StampKind { get; set; }
            public bool ShouldEmbed { get; set; }
            public string PictureName { get; set; } = string.Empty;
            public string SkipReason { get; set; } = string.Empty;
        }

        private sealed class CooperationStampCellMaps
        {
            public Dictionary<string, string> ByRoleKey { get; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> ByApprover { get; } = new(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class StampCanvasInfo
        {
            public int WidthPx { get; set; }
            public int HeightPx { get; set; }
        }

        private sealed class StampImageLayoutInfo
        {
            public string SourceImagePath { get; set; } = string.Empty;
            public bool SourceExists { get; set; }
            public int SourceWidthPx { get; set; }
            public int SourceHeightPx { get; set; }
            public int TrimLeftPx { get; set; }
            public int TrimTopPx { get; set; }
            public int TrimRightPx { get; set; }
            public int TrimBottomPx { get; set; }
            public int TrimWidthPx { get; set; }
            public int TrimHeightPx { get; set; }
        }

        private enum StampKind
        {
            Approved,
            Holding,
            Rejected
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
            var userName = User?.Identity?.Name ?? string.Empty;
            var userEmail = User?.FindFirstValue(ClaimTypes.Email) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(userId))
                return (false, false, false, false, false);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            string docStatus = string.Empty;
            string createdBy = string.Empty;

            await using (var doc = conn.CreateCommand())
            {
                doc.CommandText = @"
SELECT TOP 1
       ISNULL(CreatedBy, N'') AS CreatedBy,
       ISNULL(Status,    N'') AS Status
FROM dbo.Documents
WHERE DocId = @DocId;";
                doc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                await using var r = await doc.ExecuteReaderAsync();
                if (!await r.ReadAsync())
                    return (false, false, false, false, false);

                createdBy = r["CreatedBy"]?.ToString() ?? string.Empty;
                docStatus = r["Status"]?.ToString() ?? string.Empty;
            }

            var isCreator = string.Equals(createdBy.Trim(), userId.Trim(), StringComparison.OrdinalIgnoreCase);
            var isPendingDoc = !string.IsNullOrWhiteSpace(docStatus)
                && docStatus.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);
            var isRecalled = !string.IsNullOrWhiteSpace(docStatus)
                && docStatus.StartsWith("Recalled", StringComparison.OrdinalIgnoreCase);
            var isFinal = !string.IsNullOrWhiteSpace(docStatus)
                && !isPendingDoc
                && !string.Equals(docStatus, "Draft", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(docStatus, "Created", StringComparison.OrdinalIgnoreCase);

            int progressedApprovalCount = 0;
            await using (var ap = conn.CreateCommand())
            {
                ap.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND (
        ActedAt IS NOT NULL
        OR (
            NULLIF(LTRIM(RTRIM(ISNULL(Action, N''))), N'') IS NOT NULL
            AND ISNULL(Action, N'') <> N'Recalled'
        )
      );";
                ap.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                progressedApprovalCount = Convert.ToInt32(await ap.ExecuteScalarAsync());
            }

            int progressedCooperationCount = 0;
            await using (var co = conn.CreateCommand())
            {
                co.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentCooperations
WHERE DocId = @DocId
  AND (
        ActedAt IS NOT NULL
        OR (
            NULLIF(LTRIM(RTRIM(ISNULL(Action, N''))), N'') IS NOT NULL
            AND ISNULL(Action, N'') <> N'Recalled'
        )
      );";
                co.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                progressedCooperationCount = Convert.ToInt32(await co.ExecuteScalarAsync());
            }

            var canRecall =
                isCreator
                && isPendingDoc
                && progressedApprovalCount == 0
                && progressedCooperationCount == 0;

            bool canCooperate = false;
            await using (var coopCmd = conn.CreateCommand())
            {
                coopCmd.CommandText = @"
SELECT TOP 1 1
FROM dbo.DocumentCooperations
WHERE DocId = @id
  AND UserId = @uid
  AND ISNULL(Status, N'Pending') = N'Pending';";
                coopCmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                coopCmd.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = userId });
                var coopObj = await coopCmd.ExecuteScalarAsync();
                canCooperate = (coopObj != null && coopObj != DBNull.Value);
            }

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
                return (canRecall, false, false, false, canCooperate);

            bool canDo = false;
            await using (var cmd3 = conn.CreateCommand())
            {
                cmd3.CommandText = @"
SELECT TOP (1) 1
FROM dbo.DocumentApprovals
WHERE DocId = @id
  AND StepOrder = @step
  AND ISNULL(Status, N'Pending') LIKE N'Pending%'
  AND (
        NULLIF(LTRIM(RTRIM(UserId)), N'') = @uid
     OR NULLIF(LTRIM(RTRIM(ApproverValue)), N'') = @uid
     OR (@userName <> N'' AND NULLIF(LTRIM(RTRIM(ApproverValue)), N'') = @userName)
     OR (@userEmail <> N'' AND NULLIF(LTRIM(RTRIM(ApproverValue)), N'') = @userEmail)
  );";
                cmd3.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                cmd3.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = currentStep });
                cmd3.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 450) { Value = userId.Trim() });
                cmd3.Parameters.Add(new SqlParameter("@userName", SqlDbType.NVarChar, 256) { Value = userName.Trim() });
                cmd3.Parameters.Add(new SqlParameter("@userEmail", SqlDbType.NVarChar, 256) { Value = userEmail.Trim() });

                var mineObj = await cmd3.ExecuteScalarAsync();
                canDo = (mineObj != null && mineObj != DBNull.Value);
            }

            if (!canDo)
                return (canRecall, false, false, false, canCooperate);

            bool isCurrentFinalApprovalStep = true;
            await using (var nextCmd = conn.CreateCommand())
            {
                nextCmd.CommandText = @"
SELECT TOP 1 1
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder > @Step
ORDER BY StepOrder;";
                nextCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                nextCmd.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = currentStep });
                var nextObj = await nextCmd.ExecuteScalarAsync();
                isCurrentFinalApprovalStep = (nextObj == null || nextObj == DBNull.Value);
            }

            // 2026.06.12 Fixed: 최종 결재자는 모든 협조 완료 + 모든 이전 결재 승인 완료 전에는
            // 승인/보류/반려 버튼을 활성화하지 않는다.
            if (isCurrentFinalApprovalStep)
            {
                int remainingCooperationCount = 0;
                await using (var remainCmd = conn.CreateCommand())
                {
                    remainCmd.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentCooperations
WHERE DocId = @DocId
  AND ISNULL(Status, N'Pending') NOT IN (N'Cooperated', N'Recalled');";
                    remainCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    remainingCooperationCount = Convert.ToInt32(await remainCmd.ExecuteScalarAsync());
                }

                int previousNotApprovedCount = 0;
                await using (var prevCmd = conn.CreateCommand())
                {
                    prevCmd.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder < @Step
  AND ISNULL(Status, N'Pending') NOT IN (N'Approved');";
                    prevCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    prevCmd.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = currentStep });
                    previousNotApprovedCount = Convert.ToInt32(await prevCmd.ExecuteScalarAsync());
                }

                if (remainingCooperationCount > 0 || previousNotApprovedCount > 0)
                {
                    _log.LogInformation(
                        "DetailDX capability: final approval blocked. docId={DocId}, pendingCoop={PendingCoop}, prevNotApproved={PrevNotApproved}",
                        docId,
                        remainingCooperationCount,
                        previousNotApprovedCount);

                    return (canRecall, false, false, false, canCooperate);
                }
            }

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

        [AllowAnonymous]
        [HttpGet("DetailDXWarmup")]
        [ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
        public IActionResult DetailDXWarmup()
        {
            if (!IsLocalWarmupRequest())
            {
                _log.LogWarning(
                    "DetailDXWarmup denied remoteIp={RemoteIp} localIp={LocalIp}",
                    HttpContext?.Connection?.RemoteIpAddress?.ToString() ?? string.Empty,
                    HttpContext?.Connection?.LocalIpAddress?.ToString() ?? string.Empty);

                return NotFound();
            }

            ViewData["DisableDxAll"] = true;
            ViewData["UseDxSpreadsheet"] = true;

            // 2026.06.15 Changed: 서버 시작 1회 warm-up 호출을 허용하되 localhost 요청만 처리하도록 제한한다.
            var dxCallbackUrl = "/Doc/dx-callback";
            var dxDocumentId = "detail_warmup_" + Guid.NewGuid().ToString("N");
            var warmupPath = EnsureDetailDxWarmupWorkbookPath();
            var warmupExists = System.IO.File.Exists(warmupPath);

            ViewBag.DxCallbackUrl = dxCallbackUrl;
            ViewBag.DxDocumentId = dxDocumentId;
            ViewBag.WarmupPath = warmupPath;

            _log.LogInformation(
                "DetailDXWarmup requested documentId={DocumentId} warmupPath={WarmupPath} exists={Exists}",
                dxDocumentId,
                warmupPath,
                warmupExists);

            return View("~/Views/Doc/DetailDXWarmup.cshtml");
        }


        // 2026.06.23 Added: DocDXStamp 하위 파일만 삭제하는 안전 삭제 헬퍼 Contents App_Data Docs 원본 문서 삭제를 방지한다
        private bool TryDeleteDetailDxStampTempFile(string docId, string stampWorkPath)
        {
            try
            {
                var root = Path.GetFullPath(Path.Combine(_env.ContentRootPath, "App_Data", "DocDXStamp"));
                var fullPath = Path.GetFullPath(stampWorkPath);

                if (!fullPath.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    return false;

                var safeDocId = SafeStampFilePart(docId);
                var fileName = Path.GetFileName(fullPath);

                if (string.IsNullOrWhiteSpace(fileName) ||
                    !fileName.StartsWith(safeDocId + "_", StringComparison.OrdinalIgnoreCase))
                    return false;

                if (!System.IO.File.Exists(fullPath))
                    return false;

                System.IO.File.Delete(fullPath);

                TryDeleteEmptyDirectoriesUpToRoot(Path.GetDirectoryName(fullPath), root);

                return true;
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX stamp temp cleanup failed docId={DocId} path={Path}", docId, stampWorkPath);
                return false;
            }
        }

        // 2026.06.23 Added: 빈 임시 하위 폴더 정리 Contents 파일 삭제 후 빈 문서별 날짜 폴더를 가능한 범위에서 제거한다
        private static void TryDeleteEmptyDirectoriesUpToRoot(string? startDir, string rootDir)
        {
            if (string.IsNullOrWhiteSpace(startDir) || string.IsNullOrWhiteSpace(rootDir))
                return;

            var root = Path.GetFullPath(rootDir).TrimEnd(Path.DirectorySeparatorChar);
            var current = Path.GetFullPath(startDir);

            while (!string.Equals(current.TrimEnd(Path.DirectorySeparatorChar), root, StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    if (!Directory.Exists(current))
                        break;

                    if (Directory.EnumerateFileSystemEntries(current).Any())
                        break;

                    Directory.Delete(current, recursive: false);

                    var parent = Directory.GetParent(current);
                    if (parent == null)
                        break;

                    current = parent.FullName;
                }
                catch
                {
                    break;
                }
            }
        }

        // 2026.06.15 Added: 서버 시작 warm-up endpoint를 외부에서 호출하지 못하도록 loopback 요청만 허용한다.
        private bool IsLocalWarmupRequest()
        {
            var connection = HttpContext?.Connection;
            if (connection == null)
                return false;

            var remoteIp = connection.RemoteIpAddress;
            if (remoteIp == null)
                return false;

            if (IPAddress.IsLoopback(remoteIp))
                return true;

            var localIp = connection.LocalIpAddress;
            if (localIp != null && remoteIp.Equals(localIp))
                return true;

            return false;
        }

        // 2026.06.23 Changed: 기존 workbook 생성 방식을 제거 Contents 표준 zip 패키지로 최소 xlsx만 생성한다
        private string EnsureDetailDxWarmupWorkbookPath()
        {
            var warmupDir = Path.Combine(_env.ContentRootPath, "App_Data", "DocDXWarmup");
            Directory.CreateDirectory(warmupDir);

            var warmupPath = Path.Combine(warmupDir, "DetailDXWarmup.xlsx");
            if (System.IO.File.Exists(warmupPath))
                return warmupPath;

            lock (_detailDxWarmupWorkbookLock)
            {
                if (System.IO.File.Exists(warmupPath))
                    return warmupPath;

                var tmp = warmupPath + ".tmp";
                TryDeleteStampFile(tmp);
                WriteMinimalDetailDxWarmupWorkbook(tmp);
                System.IO.File.Move(tmp, warmupPath);
            }

            return warmupPath;
        }

        // 2026.06.23 Changed: 기존 workbook 생성 방식을 제거 Contents 표준 zip 패키지로 최소 xlsx만 생성하고 엑셀 핸들링은 수행하지 않는다
        private static void WriteMinimalDetailDxWarmupWorkbook(string path)
        {
            using var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Create);

            AddZipText(zip, "[Content_Types].xml", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
""");

            AddZipText(zip, "_rels/.rels", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""");

            AddZipText(zip, "xl/workbook.xml", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""");

            AddZipText(zip, "xl/_rels/workbook.xml.rels", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
""");

            AddZipText(zip, "xl/worksheets/sheet1.xml", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>DetailDX Warmup</t></is></c>
    </row>
  </sheetData>
</worksheet>
""");
        }

        private static void AddZipText(System.IO.Compression.ZipArchive zip, string entryName, string content)
        {
            var entry = zip.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
            writer.Write(content.TrimStart());
        }

        public sealed class ApproveDto
        {
            public string? docId { get; set; }
            public string? action { get; set; }
            public int slot { get; set; } = 1;
            public string? comment { get; set; }
            // 2026.06.23 Added: 현재 열린 DevExpress Spreadsheet 세션 상태 전달 Contents ComposeDX BatchEditDX와 동일한 SpreadsheetClientState 사용
            public SpreadsheetClientState? spreadsheetState { get; set; }
            public string? stampWorkPath { get; set; }
        }

        public sealed class BackfillStampsDto
        {
            public string? docId { get; set; }
            public SpreadsheetClientState? spreadsheetState { get; set; }
            public string? stampWorkPath { get; set; }
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

        public sealed class DetailDxStampTempCleanupRequest
        {
            public string? DocId { get; set; }
            public string? StampWorkPath { get; set; }
            public string? DxDocumentId { get; set; }
        }
    }
}
