using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security.Claims;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using System.IO;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("DocManage")]
    public class DocManageController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _env;

        public DocManageController(IConfiguration configuration, IWebHostEnvironment env)
        {
            _configuration = configuration;
            _env = env;
        }

        [HttpGet("")]
        public async Task<IActionResult> Index([FromQuery] DocManageSearchVm search)
        {
            await using var cn = new SqlConnection(GetConnectionString());
            await cn.OpenAsync();

            if (!await IsCurrentUserAdminAsync(cn))
                return Forbid();

            var normalizedSearch = NormalizeSearch(search);

            ViewData["DisableDxAll"] = false;
            ViewData["UseDxSpreadsheet"] = false;

            var model = new DocManageIndexVm
            {
                Search = normalizedSearch,
                Items = await LoadDocumentRowsAsync(cn, normalizedSearch)
            };

            return View("~/Views/Doc/DocManage.cshtml", model);
        }

        private DocManageSearchVm NormalizeSearch(DocManageSearchVm? search)
        {
            var today = DateTime.Today;

            var toDate = search?.ToDate?.Date ?? today;
            var fromDate = search?.FromDate?.Date ?? toDate.AddDays(-30);

            if (fromDate > toDate)
            {
                var tmp = fromDate;
                fromDate = toDate;
                toDate = tmp;
            }

            return new DocManageSearchVm
            {
                FromDate = fromDate,
                ToDate = toDate
            };
        }

        private async Task<List<DocManageRowVm>> LoadDocumentRowsAsync(SqlConnection cn, DocManageSearchVm search)
        {
            if (!await TableExistsAsync(cn, "dbo", "Documents"))
                return new List<DocManageRowVm>();

            var documentColumns = await LoadTableColumnsAsync(cn, "dbo", "Documents");

            var docIdCol = PickColumn(documentColumns, "DocId", "DocumentId", "DocumentNo", "DocumentCode", "Id");
            var createdAtCol = PickColumn(documentColumns, "CreatedAt", "CreatedDate", "WriteDate", "SubmittedAt", "RegDate", "InsertedAt");
            var titleCol = PickColumn(documentColumns, "TemplateTitle", "Title", "DocTitle", "DocumentTitle", "Subject");
            var statusCol = PickColumn(documentColumns, "Status", "DocStatus", "State", "ApprovalStatus");
            var createdByCol = PickColumn(documentColumns, "CreatedBy", "AuthorUserId", "WriterUserId", "UserId", "RegUserId");
            var createdByNameCol = PickColumn(documentColumns, "CreatedByName", "AuthorName", "WriterName");
            var compCdCol = PickColumn(documentColumns, "CompCd", "CompanyCode");
            var departmentIdCol = PickColumn(documentColumns, "DepartmentId", "DeptId");
            var templateCodeCol = PickColumn(documentColumns, "TemplateCode", "DocCode", "TemplateDocCode");
            var templateVersionIdCol = PickColumn(documentColumns, "TemplateVersionId", "VersionId", "DocTemplateVersionId");
            var documentFilePathCol = PickColumn(
                documentColumns,
                "OutputFilePath",
                "ExcelFilePath",
                "GeneratedExcelPath",
                "GeneratedFilePath",
                "SavedFilePath",
                "DocumentFilePath",
                "DocFilePath",
                "ResultFilePath",
                "FilePath",
                "StoragePath",
                "PhysicalPath",
                "OutputPath",
                "ExcelPath",
                "DocumentPath",
                "DocPath",
                "SavedPath"
            );

            var documentStorageCol = PickColumn(
                documentColumns,
                "ExcelStorage",
                "FileStorage",
                "Storage",
                "StorageType"
            );

            var documentBlobCol = PickColumn(
                documentColumns,
                "ExcelBlob",
                "FileBlob",
                "Blob",
                "Content",
                "FileContent",
                "OutputBlob"
            );
            var updatedAtCol = PickColumn(documentColumns, "UpdatedAt", "ModifiedAt", "LastUpdatedAt");

            if (string.IsNullOrWhiteSpace(docIdCol) || string.IsNullOrWhiteSpace(createdAtCol))
                return new List<DocManageRowVm>();

            var langCode = CultureInfo.CurrentUICulture.Name;
            if (string.IsNullOrWhiteSpace(langCode))
                langCode = "ko-KR";

            var titleExpr = SqlOptionalText("d", titleCol);
            var createdByExpr = SqlOptionalText("d", createdByCol);
            var createdByNameExpr = SqlOptionalText("d", createdByNameCol);
            var compCdExpr = SqlOptionalText("d", compCdCol);
            var departmentIdExpr = SqlOptionalText("d", departmentIdCol);
            var templateCodeExpr = SqlOptionalText("d", templateCodeCol);
            var templateVersionIdExpr = SqlOptionalText("d", templateVersionIdCol);
            var documentFilePathExpr = SqlOptionalText("d", documentFilePathCol);
            var documentStorageExpr = SqlOptionalText("d", documentStorageCol);
            var documentBlobSizeExpr = string.IsNullOrWhiteSpace(documentBlobCol)
                ? "CONVERT(bigint, 0)"
                : $@"
CASE
    WHEN {SqlColumn("d", documentBlobCol)} IS NULL THEN CONVERT(bigint, 0)
    ELSE CONVERT(bigint, DATALENGTH({SqlColumn("d", documentBlobCol)}))
END";

            var compJoin = string.IsNullOrWhiteSpace(compCdCol)
                ? ""
                : $@"
LEFT JOIN dbo.CompMasters cm
       ON cm.CompCd = {compCdExpr}";

            var deptJoin = string.IsNullOrWhiteSpace(departmentIdCol)
                ? ""
                : $@"
LEFT JOIN dbo.DepartmentMasters dm
       ON dm.Id = TRY_CONVERT(int, {departmentIdExpr})
LEFT JOIN dbo.DepartmentMasterLoc dml
       ON dml.DepartmentId = dm.Id
      AND dml.LangCode = @LangCode";

            var versionJoin = string.IsNullOrWhiteSpace(templateVersionIdCol)
                ? ""
                : $@"
LEFT JOIN dbo.DocTemplateVersion dtv
       ON dtv.Id = TRY_CONVERT(bigint, {templateVersionIdExpr})";

            var compNameExpr = string.IsNullOrWhiteSpace(compCdCol)
                ? "CAST(NULL AS nvarchar(4000))"
                : $@"
COALESCE
(
    NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(4000), cm.Name))), N''),
    NULLIF(LTRIM(RTRIM({compCdExpr})), N'')
)";
            var compTimeZoneExpr = string.IsNullOrWhiteSpace(compCdCol)
    ? "N'Asia/Seoul'"
    : @"
COALESCE
(
    NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), cm.TimeZoneId))), N''),
    N'Asia/Seoul'
)";

            var compLocaleExpr = string.IsNullOrWhiteSpace(compCdCol)
                ? $"N'{langCode.Replace("'", "''")}'"
                : $@"
COALESCE
(
    NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(20), cm.Locale))), N''),
    N'{langCode.Replace("'", "''")}'
)";

            var departmentNameExpr = string.IsNullOrWhiteSpace(departmentIdCol)
                ? "CAST(NULL AS nvarchar(4000))"
                : $@"
COALESCE
(
    NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(4000), dml.Name))), N''),
    NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(4000), dm.Name))), N''),
    NULLIF(LTRIM(RTRIM({departmentIdExpr})), N'')
)";

            var authorDisplayExpr = $@"
COALESCE
(
    NULLIF(LTRIM(RTRIM({createdByNameExpr})), N''),
    NULLIF(LTRIM(RTRIM({createdByExpr})), N'')
)";

            var templateVersionNoExpr = string.IsNullOrWhiteSpace(templateVersionIdCol)
                ? "CAST(NULL AS int)"
                : "dtv.VersionNo";

            var sql = $@"
SELECT TOP (1000)
    {SqlText("d", docIdCol)} AS DocId,
    {titleExpr} AS Title,
    {SqlOptionalText("d", statusCol)} AS Status,
    {createdByExpr} AS CreatedBy,
    {createdByNameExpr} AS CreatedByName,
    {authorDisplayExpr} AS AuthorDisplayName,
    TRY_CONVERT(datetime2(0), {SqlColumn("d", createdAtCol)}) AS CreatedAt,
    {compCdExpr} AS CompCd,
    {compNameExpr} AS CompName,
    {departmentIdExpr} AS DepartmentId,
    {departmentNameExpr} AS DepartmentName,
    {templateCodeExpr} AS TemplateCode,
    {templateVersionIdExpr} AS TemplateVersionId,
    {templateVersionNoExpr} AS TemplateVersionNo,
    {(string.IsNullOrWhiteSpace(updatedAtCol) ? "CAST(NULL AS datetime2(0))" : $"TRY_CONVERT(datetime2(0), {SqlColumn("d", updatedAtCol)})")} AS UpdatedAt,
    {documentFilePathExpr} AS DocumentFilePath,
    {documentStorageExpr} AS DocumentStorage,
    {documentBlobSizeExpr} AS DocumentBlobSize,
    {compTimeZoneExpr} AS CompTimeZoneId,
    {compLocaleExpr} AS CompLocale
FROM dbo.Documents d
{compJoin}
{deptJoin}
{versionJoin}
WHERE TRY_CONVERT(datetime2(0), {SqlColumn("d", createdAtCol)}) >= @FromDate
  AND TRY_CONVERT(datetime2(0), {SqlColumn("d", createdAtCol)}) < DATEADD(DAY, 1, @ToDate)
ORDER BY
    TRY_CONVERT(datetime2(0), {SqlColumn("d", createdAtCol)}) DESC,
    {SqlText("d", docIdCol)} DESC;";

            var list = new List<DocManageRowVm>();

            await using var cmd = cn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.Add("@FromDate", SqlDbType.DateTime2).Value = search.FromDate!.Value.Date.AddDays(-1);
            cmd.Parameters.Add("@ToDate", SqlDbType.DateTime2).Value = search.ToDate!.Value.Date.AddDays(1);
            cmd.Parameters.Add("@LangCode", SqlDbType.NVarChar, 20).Value = langCode;

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var documentFilePath = ReadString(rd, "DocumentFilePath");
                var documentStorage = ReadString(rd, "DocumentStorage");
                var documentBlobSize = ReadLong(rd, "DocumentBlobSize") ?? 0;

                var compTimeZoneId = ReadString(rd, "CompTimeZoneId");
                var compLocale = ReadString(rd, "CompLocale");

                var createdAtUtc = ReadDateTime(rd, "CreatedAt");
                var updatedAtUtc = ReadDateTime(rd, "UpdatedAt");

                var createdAtLocal = DocControllerHelper.ConvertUtcToLocal(createdAtUtc, compTimeZoneId);
                var updatedAtLocal = DocControllerHelper.ConvertUtcToLocal(updatedAtUtc, compTimeZoneId);

                if (createdAtLocal == null)
                    continue;

                if (createdAtLocal.Value.Date < search.FromDate!.Value.Date ||
                    createdAtLocal.Value.Date > search.ToDate!.Value.Date)
                    continue;

                list.Add(new DocManageRowVm
                {
                    DocId = ReadString(rd, "DocId"),
                    Title = ReadString(rd, "Title"),
                    Status = ReadString(rd, "Status"),
                    StatusCode = ReadString(rd, "Status"),
                    CreatedBy = ReadString(rd, "CreatedBy"),
                    CreatedByName = ReadString(rd, "CreatedByName"),
                    AuthorDisplayName = ReadString(rd, "AuthorDisplayName"),

                    CreatedAtUtc = DocControllerHelper.TreatAsUtc(createdAtUtc),
                    CreatedAt = createdAtLocal,
                    CreatedAtLocalText = DocControllerHelper.FormatLocalMinute(createdAtLocal),
                    CreatedAtLocalDateKey = DocControllerHelper.FormatLocalDateKey(createdAtLocal),

                    CompCd = ReadString(rd, "CompCd"),
                    CompName = ReadString(rd, "CompName"),
                    CompTimeZoneId = compTimeZoneId,
                    CompLocale = compLocale,

                    DepartmentId = ReadString(rd, "DepartmentId"),
                    DepartmentName = ReadString(rd, "DepartmentName"),
                    TemplateCode = ReadString(rd, "TemplateCode"),
                    TemplateVersionId = ReadString(rd, "TemplateVersionId"),
                    TemplateVersionNo = ReadInt(rd, "TemplateVersionNo"),

                    UpdatedAtUtc = DocControllerHelper.TreatAsUtc(updatedAtUtc),
                    UpdatedAt = updatedAtLocal,
                    UpdatedAtLocalText = DocControllerHelper.FormatLocalMinute(updatedAtLocal),

                    HasFiles = DocumentFileExists(documentStorage, documentFilePath, documentBlobSize)
                });
            }

            await ApplyApprovalLinesAsync(cn, list);
            await ApplyCooperationLinesAsync(cn, list);

            foreach (var row in list)
            {
                ApplyStatusFilterFields(row);
            }

            return list;
        }

        private async Task ApplyApprovalLinesAsync(SqlConnection cn, List<DocManageRowVm> rows)
        {
            if (rows.Count == 0)
                return;

            if (!await TableExistsAsync(cn, "dbo", "DocumentApprovals"))
                return;

            var approvalColumns = await LoadTableColumnsAsync(cn, "dbo", "DocumentApprovals");

            var docIdCol = PickColumn(approvalColumns, "DocId", "DocumentId", "DocumentNo", "DocumentCode");
            var stepOrderCol = PickColumn(approvalColumns, "StepOrder", "ApprovalOrder", "SortOrder", "StepNo");
            var roleKeyCol = PickColumn(approvalColumns, "RoleKey", "ApprovalRole", "ApprovalKey");
            var statusCol = PickColumn(approvalColumns, "Status", "ApprovalStatus", "State");
            var actionCol = PickColumn(approvalColumns, "Action", "ApprovalAction", "ResultAction", "Act");
            var userIdCol = PickColumn(approvalColumns, "UserId", "ApproverUserId");
            var approverNameCol = PickColumn(approvalColumns, "ApproverName", "ActorName", "DisplayName");
            var approverValueCol = PickColumn(approvalColumns, "ApproverValue", "Approver", "ApproverEmail");
            var approverDisplayTextCol = PickColumn(approvalColumns, "ApproverDisplayText", "DisplayText");
            var actedAtCol = PickColumn(approvalColumns, "ActedAt", "ApprovedAt", "ActionAt", "UpdatedAt");

            if (string.IsNullOrWhiteSpace(docIdCol))
                return;

            var docIds = rows
                .Select(x => x.DocId)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(1000)
                .ToList();

            if (docIds.Count == 0)
                return;

            var parameters = new List<string>();
            for (var i = 0; i < docIds.Count; i++)
                parameters.Add("@DocId" + i.ToString(CultureInfo.InvariantCulture));

            var userJoin = string.IsNullOrWhiteSpace(userIdCol)
                ? ""
                : $@"
LEFT JOIN dbo.AspNetUsers au
       ON au.Id = {SqlOptionalText("a", userIdCol)}
LEFT JOIN dbo.UserProfiles aup
       ON aup.UserId = au.Id";

            var approverNameExpr = string.IsNullOrWhiteSpace(userIdCol)
                ? $@"
COALESCE
(
    NULLIF(LTRIM(RTRIM({SqlOptionalText("a", approverDisplayTextCol)})), N''),
    NULLIF(LTRIM(RTRIM({SqlOptionalText("a", approverNameCol)})), N''),
    NULLIF(LTRIM(RTRIM({SqlOptionalText("a", approverValueCol)})), N'')
)"
                : $@"
COALESCE
(
    NULLIF(LTRIM(RTRIM({SqlOptionalText("a", approverDisplayTextCol)})), N''),
    NULLIF(LTRIM(RTRIM(aup.DisplayName)), N''),
    NULLIF(LTRIM(RTRIM(au.UserName)), N''),
    NULLIF(LTRIM(RTRIM({SqlOptionalText("a", approverNameCol)})), N''),
    NULLIF(LTRIM(RTRIM({SqlOptionalText("a", approverValueCol)})), N'')
)";

            var sql = $@"
SELECT
    {SqlText("a", docIdCol)} AS DocId,
    {(string.IsNullOrWhiteSpace(stepOrderCol) ? "0" : $"TRY_CONVERT(int, {SqlColumn("a", stepOrderCol)})")} AS StepOrder,
    {SqlOptionalText("a", roleKeyCol)} AS RoleKey,
    {SqlOptionalText("a", statusCol)} AS Status,
    {SqlOptionalText("a", actionCol)} AS Action,
    {SqlOptionalText("a", userIdCol)} AS UserId,
    {approverNameExpr} AS ApproverName,
    {(string.IsNullOrWhiteSpace(actedAtCol) ? "CAST(NULL AS datetime2(0))" : $"TRY_CONVERT(datetime2(0), {SqlColumn("a", actedAtCol)})")} AS ActedAt
FROM dbo.DocumentApprovals a
{userJoin}
WHERE {SqlText("a", docIdCol)} IN ({string.Join(", ", parameters)})
ORDER BY
    {SqlText("a", docIdCol)} ASC,
    {(string.IsNullOrWhiteSpace(stepOrderCol) ? "0" : $"TRY_CONVERT(int, {SqlColumn("a", stepOrderCol)})")} ASC,
    {SqlOptionalText("a", roleKeyCol)} ASC;";

            var map = rows
                .Where(x => !string.IsNullOrWhiteSpace(x.DocId))
                .ToDictionary(x => x.DocId!, x => x, StringComparer.OrdinalIgnoreCase);

            await using var cmd = cn.CreateCommand();
            cmd.CommandText = sql;

            for (var i = 0; i < docIds.Count; i++)
                cmd.Parameters.Add(parameters[i], SqlDbType.NVarChar, 100).Value = docIds[i]!;

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var docId = ReadString(rd, "DocId");
                if (string.IsNullOrWhiteSpace(docId))
                    continue;

                if (!map.TryGetValue(docId, out var row))
                    continue;

                var stepOrder = ReadInt(rd, "StepOrder") ?? 0;
                var roleKey = ReadString(rd, "RoleKey");

                row.Approvals.Add(new DocManageApprovalVm
                {
                    StepOrder = stepOrder,
                    RoleKey = string.IsNullOrWhiteSpace(roleKey)
                        ? (stepOrder > 0 ? "A" + stepOrder.ToString(CultureInfo.InvariantCulture) : "")
                        : roleKey,
                    Status = ReadString(rd, "Status"),
                    Action = ReadString(rd, "Action"),
                    UserId = ReadString(rd, "UserId"),
                    ApproverName = ReadString(rd, "ApproverName"),
                    ActedAt = ReadDateTime(rd, "ActedAt")
                });
            }

            foreach (var row in rows)
            {
                ApplyBoardStatusFields(row);
            }
        }

        private async Task ApplyCooperationLinesAsync(SqlConnection cn, List<DocManageRowVm> rows)
        {
            if (rows.Count == 0)
                return;

            if (!await TableExistsAsync(cn, "dbo", "DocumentCooperations"))
                return;

            var coopColumns = await LoadTableColumnsAsync(cn, "dbo", "DocumentCooperations");

            var docIdCol = PickColumn(coopColumns, "DocId", "DocumentId", "DocumentNo", "DocumentCode");
            var lineTypeCol = PickColumn(coopColumns, "LineType");
            var roleKeyCol = PickColumn(coopColumns, "RoleKey", "CooperationRole", "CooperationKey");
            var statusCol = PickColumn(coopColumns, "Status", "CooperationStatus", "State");
            var actionCol = PickColumn(coopColumns, "Action", "CooperationAction", "ResultAction", "Act");
            var actorNameCol = PickColumn(coopColumns, "ActorName", "ApproverName", "DisplayName");
            var approverDisplayTextCol = PickColumn(coopColumns, "ApproverDisplayText", "DisplayText");
            var approverValueCol = PickColumn(coopColumns, "ApproverValue", "Approver", "ApproverEmail");
            var userIdCol = PickColumn(coopColumns, "UserId", "ApproverUserId");
            var actedAtCol = PickColumn(coopColumns, "ActedAt", "ActionAt", "UpdatedAt");

            if (string.IsNullOrWhiteSpace(docIdCol))
                return;

            var docIds = rows
                .Select(x => x.DocId)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(1000)
                .ToList();

            if (docIds.Count == 0)
                return;

            var parameters = new List<string>();
            for (var i = 0; i < docIds.Count; i++)
                parameters.Add("@CoopDocId" + i.ToString(CultureInfo.InvariantCulture));

            var userJoin = string.IsNullOrWhiteSpace(userIdCol)
                ? ""
                : $@"
LEFT JOIN dbo.AspNetUsers au
       ON au.Id = {SqlOptionalText("c", userIdCol)}
LEFT JOIN dbo.UserProfiles aup
       ON aup.UserId = au.Id";

            var coopNameExpr = string.IsNullOrWhiteSpace(userIdCol)
                ? $@"
COALESCE
(
    NULLIF(LTRIM(RTRIM({SqlOptionalText("c", approverDisplayTextCol)})), N''),
    NULLIF(LTRIM(RTRIM({SqlOptionalText("c", actorNameCol)})), N''),
    NULLIF(LTRIM(RTRIM({SqlOptionalText("c", approverValueCol)})), N'')
)"
                : $@"
COALESCE
(
    NULLIF(LTRIM(RTRIM({SqlOptionalText("c", approverDisplayTextCol)})), N''),
    NULLIF(LTRIM(RTRIM(aup.DisplayName)), N''),
    NULLIF(LTRIM(RTRIM(au.UserName)), N''),
    NULLIF(LTRIM(RTRIM({SqlOptionalText("c", actorNameCol)})), N''),
    NULLIF(LTRIM(RTRIM({SqlOptionalText("c", approverValueCol)})), N'')
)";

            var lineTypeFilter = string.IsNullOrWhiteSpace(lineTypeCol)
                ? ""
                : $"AND UPPER({SqlOptionalText("c", lineTypeCol)}) = N'COOPERATION'";

            var sql = $@"
SELECT
    {SqlText("c", docIdCol)} AS DocId,
    {SqlOptionalText("c", roleKeyCol)} AS RoleKey,
    {SqlOptionalText("c", statusCol)} AS Status,
    {SqlOptionalText("c", actionCol)} AS Action,
    {SqlOptionalText("c", userIdCol)} AS UserId,
    {coopNameExpr} AS CoopName,
    {(string.IsNullOrWhiteSpace(actedAtCol) ? "CAST(NULL AS datetime2(0))" : $"TRY_CONVERT(datetime2(0), {SqlColumn("c", actedAtCol)})")} AS ActedAt
FROM dbo.DocumentCooperations c
{userJoin}
WHERE {SqlText("c", docIdCol)} IN ({string.Join(", ", parameters)})
{lineTypeFilter}
ORDER BY
    {SqlText("c", docIdCol)} ASC,
    {SqlOptionalText("c", roleKeyCol)} ASC;";

            var map = rows
                .Where(x => !string.IsNullOrWhiteSpace(x.DocId))
                .ToDictionary(x => x.DocId!, x => x, StringComparer.OrdinalIgnoreCase);

            await using var cmd = cn.CreateCommand();
            cmd.CommandText = sql;

            for (var i = 0; i < docIds.Count; i++)
                cmd.Parameters.Add(parameters[i], SqlDbType.NVarChar, 100).Value = docIds[i]!;

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var docId = ReadString(rd, "DocId");
                if (string.IsNullOrWhiteSpace(docId))
                    continue;

                if (!map.TryGetValue(docId, out var row))
                    continue;

                var roleKey = ReadString(rd, "RoleKey");

                row.Cooperations.Add(new DocManageCooperationVm
                {
                    RoleKey = roleKey,
                    StepNo = ParseRoleNo(roleKey),
                    Status = ReadString(rd, "Status"),
                    Action = ReadString(rd, "Action"),
                    UserId = ReadString(rd, "UserId"),
                    CoopName = ReadString(rd, "CoopName"),
                    ActedAt = ReadDateTime(rd, "ActedAt")
                });
            }

            foreach (var row in rows)
            {
                ApplyBoardCooperationFields(row);
            }
        }

        private static void ApplyBoardStatusFields(DocManageRowVm row)
        {
            row.StatusCode = row.Status;
            row.TotalApprovers = row.Approvals.Count;
            row.CompletedApprovers = row.Approvals.Count(x => IsApprovedStatus(x.Status) || IsApprovedStatus(x.Action));
            row.ApprovalSummary = BuildApprovalSummary(row.Approvals);

            row.ApprovalSteps = row.Approvals
                .OrderBy(x => x.StepOrder)
                .ThenBy(x => x.RoleKey)
                .Select(x => new DocManageApprovalStepVm
                {
                    StepOrder = x.StepOrder,
                    RoleKey = x.RoleKey,
                    Status = NormalizeBoardStepStatus(x.Status, x.Action),
                    Action = NormalizeBoardAction(x.Status, x.Action),
                    ApproverName = x.ApproverName
                })
                .ToList();

            row.ResultSummary = ResolveResultSummary(row.Approvals);
        }

        private static void ApplyBoardCooperationFields(DocManageRowVm row)
        {
            var ordered = row.Cooperations
                .OrderBy(x => x.StepNo <= 0 ? int.MaxValue : x.StepNo)
                .ThenBy(x => x.RoleKey)
                .ToList();

            row.CoopTotalSteps = ordered.Count;

            if (ordered.Count == 0)
                return;

            row.CoopDoneKeys = JoinStepKeys(ordered.Where(x => IsCooperatedStatus(x.Status) || IsCooperatedStatus(x.Action)));
            row.CoopRejectedKeys = JoinStepKeys(ordered.Where(x => IsRejectedStatus(x.Status) || IsRejectedStatus(x.Action)));
            row.CoopHoldKeys = JoinStepKeys(ordered.Where(x => IsHoldStatus(x.Status) || IsHoldStatus(x.Action)));
            row.CoopRecalledKeys = JoinStepKeys(ordered.Where(x => IsRecalledStatus(x.Status) || IsRecalledStatus(x.Action)));

            var pending = ordered.FirstOrDefault(x =>
                !IsCooperatedStatus(x.Status)
                && !IsCooperatedStatus(x.Action)
                && !IsRejectedStatus(x.Status)
                && !IsRejectedStatus(x.Action)
                && !IsHoldStatus(x.Status)
                && !IsHoldStatus(x.Action)
                && !IsRecalledStatus(x.Status)
                && !IsRecalledStatus(x.Action));

            if (pending != null)
            {
                row.CoopPendingName = FirstNotEmpty(pending.CoopName, pending.RoleKey);
                row.CoopPendingPosition = null;
            }
        }

        private static string? ResolveResultSummary(List<DocManageApprovalVm> approvals)
        {
            if (approvals.Count == 0)
                return null;

            var ordered = approvals
                .OrderBy(x => x.StepOrder)
                .ThenBy(x => x.RoleKey)
                .ToList();

            var rejected = ordered.FirstOrDefault(x => IsRejectedStatus(x.Status) || IsRejectedStatus(x.Action));
            if (rejected != null)
                return FirstNotEmpty(rejected.ApproverName, rejected.RoleKey);

            var held = ordered.FirstOrDefault(x => IsHoldStatus(x.Status) || IsHoldStatus(x.Action));
            if (held != null)
                return FirstNotEmpty(held.ApproverName, held.RoleKey);

            var pending = ordered.FirstOrDefault(x =>
                !IsApprovedStatus(x.Status)
                && !IsApprovedStatus(x.Action)
                && !IsRejectedStatus(x.Status)
                && !IsRejectedStatus(x.Action)
                && !IsRecalledStatus(x.Status)
                && !IsRecalledStatus(x.Action)
                && !IsHoldStatus(x.Status)
                && !IsHoldStatus(x.Action));

            if (pending != null)
                return FirstNotEmpty(pending.ApproverName, pending.RoleKey);

            var lastApproved = ordered.LastOrDefault(x => IsApprovedStatus(x.Status) || IsApprovedStatus(x.Action));
            return lastApproved == null
                ? null
                : FirstNotEmpty(lastApproved.ApproverName, lastApproved.RoleKey);
        }

        private static string BuildApprovalSummary(List<DocManageApprovalVm> approvals)
        {
            if (approvals.Count == 0)
                return "";

            var approved = approvals.Count(x => IsApprovedStatus(x.Status) || IsApprovedStatus(x.Action));
            var total = approvals.Count;

            return approved.ToString(CultureInfo.InvariantCulture)
                + "/"
                + total.ToString(CultureInfo.InvariantCulture);
        }

        private static string NormalizeBoardStepStatus(string? status, string? action)
        {
            if (IsApprovedStatus(action) || IsApprovedStatus(status))
                return "Approved";

            if (IsRejectedStatus(action) || IsRejectedStatus(status))
                return "Rejected";

            if (IsRecalledStatus(action) || IsRecalledStatus(status))
                return "Recalled";

            if (IsHoldStatus(action) || IsHoldStatus(status))
                return "OnHold";

            return "Pending";
        }

        private static string? NormalizeBoardAction(string? status, string? action)
        {
            if (IsApprovedStatus(action) || IsApprovedStatus(status))
                return "Approve";

            if (IsRejectedStatus(action) || IsRejectedStatus(status))
                return "Reject";

            if (IsRecalledStatus(action) || IsRecalledStatus(status))
                return "Recall";

            if (IsHoldStatus(action) || IsHoldStatus(status))
                return "Hold";

            return null;
        }

        private static string JoinStepKeys(IEnumerable<DocManageCooperationVm> rows)
        {
            return string.Join(",", rows
                .Select(x => x.StepNo)
                .Where(x => x > 0)
                .Distinct()
                .OrderBy(x => x)
                .Select(x => x.ToString(CultureInfo.InvariantCulture)));
        }

        private static int ParseRoleNo(string? roleKey)
        {
            if (string.IsNullOrWhiteSpace(roleKey))
                return 0;

            var digits = new string(roleKey.Where(char.IsDigit).ToArray());
            return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)
                ? n
                : 0;
        }

        private static bool IsApprovedStatus(string? value)
        {
            var s = NormalizeStatusText(value);
            return s == "APPROVE" || s.StartsWith("APPROVED", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCooperatedStatus(string? value)
        {
            var s = NormalizeStatusText(value);
            return s == "COOPERATE"
                || s.StartsWith("COOPERATED", StringComparison.OrdinalIgnoreCase)
                || s == "APPROVE"
                || s.StartsWith("APPROVED", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRejectedStatus(string? value)
        {
            var s = NormalizeStatusText(value);
            return s == "REJECT" || s.StartsWith("REJECTED", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRecalledStatus(string? value)
        {
            var s = NormalizeStatusText(value);
            return s == "RECALL" || s.StartsWith("RECALLED", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHoldStatus(string? value)
        {
            var s = NormalizeStatusText(value);
            return s == "HOLD"
                || s.StartsWith("ONHOLD", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith("ON HOLD", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith("PENDINGHOLD", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeStatusText(string? value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? ""
                : value.Trim().ToUpperInvariant();
        }

        private static string? FirstNotEmpty(params string?[] values)
        {
            return values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
        }

        private static void ApplyStatusFilterFields(DocManageRowVm row)
        {
            row.StatusFilterKeys.Clear();
            row.StatusFilterItems.Clear();
            row.StatusFilterText = "";

            var isRecalled = IsRecalledStatus(row.Status)
                || row.Approvals.Any(x => IsRecalledStatus(x.Status) || IsRecalledStatus(x.Action))
                || row.Cooperations.Any(x => IsRecalledStatus(x.Status) || IsRecalledStatus(x.Action));

            var isRejected = IsRejectedStatus(row.Status)
                || row.Approvals.Any(x => IsRejectedStatus(x.Status) || IsRejectedStatus(x.Action))
                || row.Cooperations.Any(x => IsRejectedStatus(x.Status) || IsRejectedStatus(x.Action));

            var isHeld = IsHoldStatus(row.Status)
                || row.Approvals.Any(x => IsHoldStatus(x.Status) || IsHoldStatus(x.Action))
                || row.Cooperations.Any(x => IsHoldStatus(x.Status) || IsHoldStatus(x.Action));

            var isApproved = IsRowApproved(row);

            if (isRecalled)
            {
                AddStatusFilterItem(row, "STATUS:RECALLED", "status", "Recalled", null, 0, "040");
            }
            else if (isRejected)
            {
                AddStatusFilterItem(row, "STATUS:REJECTED", "status", "Rejected", null, 0, "020");
            }
            else if (isHeld)
            {
                AddStatusFilterItem(row, "STATUS:ONHOLD", "status", "OnHold", null, 0, "030");
            }
            else if (isApproved)
            {
                AddStatusFilterItem(row, "STATUS:APPROVED", "status", "Approved", null, 0, "010");
            }
            else
            {
                var pendingApproval = row.Approvals
                    .OrderBy(x => x.StepOrder <= 0 ? int.MaxValue : x.StepOrder)
                    .ThenBy(x => x.RoleKey)
                    .FirstOrDefault(IsPendingApproval);

                if (pendingApproval != null)
                {
                    var pendingName = FirstNotEmpty(pendingApproval.ApproverName, pendingApproval.RoleKey);
                    var key = BuildPersonFilterKey("APPROVAL_PENDING", pendingApproval.UserId, pendingName);

                    if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(pendingName))
                    {
                        AddStatusFilterItem(row, key, "approvalPending", null, pendingName, 1, pendingName);
                    }
                }

                foreach (var pendingCoop in row.Cooperations
                    .Where(IsPendingCooperation)
                    .OrderBy(x => x.CoopName)
                    .ThenBy(x => x.RoleKey))
                {
                    var pendingName = FirstNotEmpty(pendingCoop.CoopName, pendingCoop.RoleKey);
                    var key = BuildPersonFilterKey("COOP_PENDING", pendingCoop.UserId, pendingName);

                    if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(pendingName))
                    {
                        AddStatusFilterItem(row, key, "coopPending", null, pendingName, 2, pendingName);
                    }
                }
            }

            row.StatusFilterText = row.StatusFilterKeys.Count == 0
                ? ""
                : "|" + string.Join("|", row.StatusFilterKeys) + "|";
        }

        private static bool IsRowApproved(DocManageRowVm row)
        {
            if (IsApprovedStatus(row.Status))
                return true;

            if (row.Approvals.Count == 0)
                return false;

            var allApprovalsApproved = row.Approvals.All(x => IsApprovedStatus(x.Status) || IsApprovedStatus(x.Action));
            var allCooperationsDone = row.Cooperations.Count == 0
                || row.Cooperations.All(x => IsCooperatedStatus(x.Status) || IsCooperatedStatus(x.Action));

            return allApprovalsApproved && allCooperationsDone;
        }

        private static bool IsPendingApproval(DocManageApprovalVm row)
        {
            return !IsApprovedStatus(row.Status)
                && !IsApprovedStatus(row.Action)
                && !IsRejectedStatus(row.Status)
                && !IsRejectedStatus(row.Action)
                && !IsRecalledStatus(row.Status)
                && !IsRecalledStatus(row.Action)
                && !IsHoldStatus(row.Status)
                && !IsHoldStatus(row.Action);
        }

        private static bool IsPendingCooperation(DocManageCooperationVm row)
        {
            return !IsCooperatedStatus(row.Status)
                && !IsCooperatedStatus(row.Action)
                && !IsRejectedStatus(row.Status)
                && !IsRejectedStatus(row.Action)
                && !IsRecalledStatus(row.Status)
                && !IsRecalledStatus(row.Action)
                && !IsHoldStatus(row.Status)
                && !IsHoldStatus(row.Action);
        }

        private static void AddStatusFilterItem(
            DocManageRowVm row,
            string key,
            string kind,
            string? code,
            string? name,
            int sortGroup,
            string? sortText)
        {
            if (string.IsNullOrWhiteSpace(key))
                return;

            if (row.StatusFilterKeys.Any(x => string.Equals(x, key, StringComparison.OrdinalIgnoreCase)))
                return;

            row.StatusFilterKeys.Add(key);
            row.StatusFilterItems.Add(new DocManageStatusFilterItemVm
            {
                Key = key,
                Kind = kind,
                Code = code,
                Name = name,
                SortGroup = sortGroup,
                SortText = sortText
            });
        }

        private static string? BuildPersonFilterKey(string prefix, string? userId, string? name)
        {
            var value = FirstNotEmpty(userId, name);
            if (string.IsNullOrWhiteSpace(value))
                return null;

            return prefix + ":" + NormalizeFilterKeyPart(value);
        }

        private static string NormalizeFilterKeyPart(string value)
        {
            return Uri.EscapeDataString(value.Trim().ToUpperInvariant());
        }

        private bool DocumentFileExists(string? storage, string? dbPath, long blobSize)
        {
            if (string.Equals(storage, "Db", StringComparison.OrdinalIgnoreCase))
                return blobSize > 0;

            if (string.IsNullOrWhiteSpace(dbPath) && blobSize > 0)
                return true;

            var absolutePath = ToSafeAbsolutePath(dbPath);
            return !string.IsNullOrWhiteSpace(absolutePath) && System.IO.File.Exists(absolutePath);
        }

        private string? ToSafeAbsolutePath(string? dbPath)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
                return null;

            var normalized = NormalizeDbPath(dbPath);
            if (string.IsNullOrWhiteSpace(normalized))
                return null;

            var absolute = Path.IsPathRooted(normalized)
                ? normalized
                : Path.Combine(_env.ContentRootPath, normalized);

            var full = Path.GetFullPath(absolute);
            var contentRoot = Path.GetFullPath(_env.ContentRootPath);

            if (!full.StartsWith(contentRoot, StringComparison.OrdinalIgnoreCase))
                return null;

            return full;
        }

        private static string? NormalizeDbPath(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;

            return path
                .Trim()
                .Replace('₩', '\\')
                .Replace('￦', '\\')
                .Replace('/', '\\');
        }

        private async Task<bool> IsCurrentUserAdminAsync(SqlConnection cn)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrWhiteSpace(userId))
                return false;

            await using var cmd = cn.CreateCommand();
            cmd.CommandText = @"
SELECT CASE
       WHEN EXISTS
       (
           SELECT 1
           FROM dbo.AspNetUsers u
           LEFT JOIN dbo.UserProfiles p
                  ON p.UserId = u.Id
           WHERE u.Id = @UserId
             AND
             (
                    ISNULL(u.IsAdmin, 0) IN (1, 2)
                 OR ISNULL(p.IsAdmin, 0) IN (1, 2)
             )
       )
       THEN 1 ELSE 0 END;";

            cmd.Parameters.Add("@UserId", SqlDbType.NVarChar, 450).Value = userId;

            var result = await cmd.ExecuteScalarAsync();
            return Convert.ToInt32(result, CultureInfo.InvariantCulture) == 1;
        }

        private async Task<bool> TableExistsAsync(SqlConnection cn, string schema, string table)
        {
            await using var cmd = cn.CreateCommand();
            cmd.CommandText = @"
SELECT CASE
       WHEN EXISTS
       (
           SELECT 1
           FROM INFORMATION_SCHEMA.TABLES
           WHERE TABLE_SCHEMA = @Schema
             AND TABLE_NAME = @Table
       )
       THEN 1 ELSE 0 END;";

            cmd.Parameters.Add("@Schema", SqlDbType.NVarChar, 128).Value = schema;
            cmd.Parameters.Add("@Table", SqlDbType.NVarChar, 128).Value = table;

            var result = await cmd.ExecuteScalarAsync();
            return Convert.ToInt32(result, CultureInfo.InvariantCulture) == 1;
        }

        private async Task<HashSet<string>> LoadTableColumnsAsync(SqlConnection cn, string schema, string table)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            await using var cmd = cn.CreateCommand();
            cmd.CommandText = @"
SELECT COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = @Schema
  AND TABLE_NAME = @Table;";

            cmd.Parameters.Add("@Schema", SqlDbType.NVarChar, 128).Value = schema;
            cmd.Parameters.Add("@Table", SqlDbType.NVarChar, 128).Value = table;

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var name = ReadString(rd, "COLUMN_NAME");
                if (!string.IsNullOrWhiteSpace(name))
                    set.Add(name);
            }

            return set;
        }

        private static string? PickColumn(HashSet<string> columns, params string[] candidates)
        {
            foreach (var candidate in candidates)
            {
                if (columns.Contains(candidate))
                    return candidate;
            }

            return null;
        }

        private static string SqlColumn(string alias, string column)
        {
            return alias + "." + QuoteSqlName(column);
        }

        private static string SqlText(string alias, string column)
        {
            return "CONVERT(nvarchar(4000), " + SqlColumn(alias, column) + ")";
        }

        private static string SqlOptionalText(string alias, string? column)
        {
            if (string.IsNullOrWhiteSpace(column))
                return "CAST(NULL AS nvarchar(4000))";

            return SqlText(alias, column);
        }

        private static string QuoteSqlName(string value)
        {
            return "[" + value.Replace("]", "]]", StringComparison.Ordinal) + "]";
        }

        private string GetConnectionString()
        {
            var cs = _configuration.GetConnectionString("DefaultConnection");
            if (string.IsNullOrWhiteSpace(cs))
                throw new InvalidOperationException("DefaultConnection is not configured.");

            return cs;
        }

        private static string? ReadString(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            if (rd.IsDBNull(ordinal))
                return null;

            return Convert.ToString(rd.GetValue(ordinal), CultureInfo.InvariantCulture);
        }

        private static int? ReadInt(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            if (rd.IsDBNull(ordinal))
                return null;

            return Convert.ToInt32(rd.GetValue(ordinal), CultureInfo.InvariantCulture);
        }

        private static long? ReadLong(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            if (rd.IsDBNull(ordinal))
                return null;

            return Convert.ToInt64(rd.GetValue(ordinal), CultureInfo.InvariantCulture);
        }

        private static DateTime? ReadDateTime(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            if (rd.IsDBNull(ordinal))
                return null;

            return Convert.ToDateTime(rd.GetValue(ordinal), CultureInfo.InvariantCulture);
        }

        private static bool ReadBool(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            if (rd.IsDBNull(ordinal))
                return false;

            return Convert.ToBoolean(rd.GetValue(ordinal), CultureInfo.InvariantCulture);
        }

        public sealed class DocManageIndexVm
        {
            public DocManageSearchVm Search { get; set; } = new();
            public List<DocManageRowVm> Items { get; set; } = new();
        }

        public sealed class DocManageSearchVm
        {
            public DateTime? FromDate { get; set; }
            public DateTime? ToDate { get; set; }
        }

        public sealed class DocManageRowVm
        {
            public string? DocId { get; set; }
            public string? Title { get; set; }
            public string? Status { get; set; }

            [JsonPropertyName("statusCode")]
            public string? StatusCode { get; set; }

            public string? CreatedBy { get; set; }
            public string? CreatedByName { get; set; }
            public string? AuthorDisplayName { get; set; }
            public DateTime? CreatedAt { get; set; }

            public string? CompCd { get; set; }
            public string? CompName { get; set; }

            public string? DepartmentId { get; set; }
            public string? DepartmentName { get; set; }

            public string? TemplateCode { get; set; }
            public string? TemplateVersionId { get; set; }
            public int? TemplateVersionNo { get; set; }

            public DateTime? UpdatedAt { get; set; }
            public string? ApprovalSummary { get; set; }
            public List<DocManageApprovalVm> Approvals { get; set; } = new();
            public List<DocManageCooperationVm> Cooperations { get; set; } = new();

            [JsonPropertyName("totalApprovers")]
            public int TotalApprovers { get; set; }

            [JsonPropertyName("completedApprovers")]
            public int CompletedApprovers { get; set; }

            [JsonPropertyName("resultSummary")]
            public string? ResultSummary { get; set; }

            [JsonPropertyName("approvalSteps")]
            public List<DocManageApprovalStepVm> ApprovalSteps { get; set; } = new();

            [JsonPropertyName("coopTotalSteps")]
            public int CoopTotalSteps { get; set; }

            [JsonPropertyName("coopDoneKeys")]
            public string? CoopDoneKeys { get; set; }

            [JsonPropertyName("coopRejectedKeys")]
            public string? CoopRejectedKeys { get; set; }

            [JsonPropertyName("coopHoldKeys")]
            public string? CoopHoldKeys { get; set; }

            [JsonPropertyName("coopRecalledKeys")]
            public string? CoopRecalledKeys { get; set; }

            [JsonPropertyName("coopPendingName")]
            public string? CoopPendingName { get; set; }

            [JsonPropertyName("coopPendingPosition")]
            public string? CoopPendingPosition { get; set; }

            [JsonPropertyName("statusFilterKeys")]
            public List<string> StatusFilterKeys { get; set; } = new();

            [JsonPropertyName("statusFilterText")]
            public string? StatusFilterText { get; set; }

            [JsonPropertyName("statusFilterItems")]
            public List<DocManageStatusFilterItemVm> StatusFilterItems { get; set; } = new();

            public bool HasFiles { get; set; }
            public DateTime? CreatedAtUtc { get; set; }
            public string? CreatedAtLocalText { get; set; }
            public string? CreatedAtLocalDateKey { get; set; }

            public string? CompTimeZoneId { get; set; }
            public string? CompLocale { get; set; }

            public DateTime? UpdatedAtUtc { get; set; }
            public string? UpdatedAtLocalText { get; set; }
        }

        public sealed class DocManageStatusFilterItemVm
        {
            [JsonPropertyName("key")]
            public string? Key { get; set; }

            [JsonPropertyName("kind")]
            public string? Kind { get; set; }

            [JsonPropertyName("code")]
            public string? Code { get; set; }

            [JsonPropertyName("name")]
            public string? Name { get; set; }

            [JsonPropertyName("sortGroup")]
            public int SortGroup { get; set; }

            [JsonPropertyName("sortText")]
            public string? SortText { get; set; }
        }

        public sealed class DocManageApprovalVm
        {
            public int StepOrder { get; set; }
            public string? RoleKey { get; set; }
            public string? Status { get; set; }
            public string? Action { get; set; }
            public string? UserId { get; set; }
            public string? ApproverName { get; set; }
            public DateTime? ActedAt { get; set; }
        }

        public sealed class DocManageApprovalStepVm
        {
            [JsonPropertyName("stepOrder")]
            public int StepOrder { get; set; }

            [JsonPropertyName("roleKey")]
            public string? RoleKey { get; set; }

            [JsonPropertyName("status")]
            public string? Status { get; set; }

            [JsonPropertyName("action")]
            public string? Action { get; set; }

            [JsonPropertyName("approverName")]
            public string? ApproverName { get; set; }
        }

        public sealed class DocManageCooperationVm
        {
            public string? RoleKey { get; set; }
            public int StepNo { get; set; }
            public string? Status { get; set; }
            public string? Action { get; set; }
            public string? UserId { get; set; }
            public string? CoopName { get; set; }
            public DateTime? ActedAt { get; set; }
        }
    }
}