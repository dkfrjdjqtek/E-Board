// 2026.06.08 Changed: 템플릿 관리 수정일을 CompMasters.TimeZoneId 기준 로컬 시간으로 변환하고 SaveUseState 저장 시간을 UTC로 통일
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("DocTL/DocTemplateManage")]
    public class DocTemplateManageController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _env;
        private readonly TemplateVersionPrepareService _templateVersionPrepareService;

        public DocTemplateManageController(
            IConfiguration configuration,
            IWebHostEnvironment env,
            TemplateVersionPrepareService templateVersionPrepareService)
        {
            _configuration = configuration;
            _env = env;
            _templateVersionPrepareService = templateVersionPrepareService;
        }

        [HttpGet("")]
        public async Task<IActionResult> Index([FromQuery] DocTemplateManageSearchVm search)
        {
            await using var cn = new SqlConnection(GetConnectionString());
            await cn.OpenAsync();

            if (!await IsCurrentUserAdminAsync(cn))
                return Forbid();

            ViewData["DisableDxAll"] = false;
            ViewData["UseDxSpreadsheet"] = false;

            var model = new DocTemplateManageIndexVm
            {
                Search = search ?? new DocTemplateManageSearchVm(),
                Items = await LoadLatestTemplateRowsAsync(cn)
            };

            return View("~/Views/DocTL/DocTemplateManage.cshtml", model);
        }

        [HttpGet("download/{versionId:long}")]
        public async Task<IActionResult> Download(long versionId)
        {
            await using var cn = new SqlConnection(GetConnectionString());
            await cn.OpenAsync();

            if (!await IsCurrentUserAdminAsync(cn))
                return Forbid();

            var file = await LoadDownloadFileAsync(cn, versionId);
            if (file == null)
                return NotFound();

            if (string.Equals(file.Storage, "Db", StringComparison.OrdinalIgnoreCase) && file.Blob != null && file.Blob.Length > 0)
            {
                var dbFileName = string.IsNullOrWhiteSpace(file.FileName) ? "template.xlsx" : file.FileName;
                var dbContentType = string.IsNullOrWhiteSpace(file.ContentType)
                    ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    : file.ContentType;

                return File(file.Blob, dbContentType, dbFileName);
            }

            var absolutePath = ToSafeAbsolutePath(file.FilePath);
            if (string.IsNullOrWhiteSpace(absolutePath) || !System.IO.File.Exists(absolutePath))
                return NotFound();

            var fileName = string.IsNullOrWhiteSpace(file.FileName)
                ? Path.GetFileName(absolutePath)
                : file.FileName;

            var contentType = string.IsNullOrWhiteSpace(file.ContentType)
                ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                : file.ContentType;

            return PhysicalFile(absolutePath, contentType, fileName);
        }

        [HttpPost("save-use-state")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SaveUseState([FromBody] SaveUseStateRequest request)
        {
            await using var cn = new SqlConnection(GetConnectionString());
            await cn.OpenAsync();

            if (!await IsCurrentUserAdminAsync(cn))
                return Forbid();

            if (request?.Items == null || request.Items.Count == 0)
                return BadRequest();

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? User.Identity?.Name ?? "system";

            await using var tx = await cn.BeginTransactionAsync();

            try
            {
                foreach (var item in request.Items.GroupBy(x => x.TemplateId).Select(x => x.Last()))
                {
                    await using var cmd = cn.CreateCommand();
                    cmd.Transaction = (SqlTransaction)tx;
                    cmd.CommandText = @"
UPDATE dbo.DocTemplateMaster
   SET IsActive = @IsActive,
       UpdatedBy = @UpdatedBy,
       UpdatedAt = SYSUTCDATETIME()
 WHERE Id = @TemplateId;";

                    cmd.Parameters.Add("@IsActive", SqlDbType.Bit).Value = item.IsActive;
                    cmd.Parameters.Add("@UpdatedBy", SqlDbType.NVarChar, 100).Value = userId;
                    cmd.Parameters.Add("@TemplateId", SqlDbType.Int).Value = item.TemplateId;

                    await cmd.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();

                return Json(new
                {
                    ok = true,
                    messageKey = "DTM_Message_SaveCompleted"
                });
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }
        }

        // 2026.06.15 Added: 기존 템플릿 버전 파일을 백업한 뒤 현재 템플릿 준비 규칙으로 보호와 표시 메타 값을 일괄 보정한다.
        [HttpPost("backfill-visual-metrics")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> BackfillVisualMetrics()
        {
            await using var cn = new SqlConnection(GetConnectionString());
            await cn.OpenAsync();

            if (!await IsCurrentUserAdminAsync(cn))
                return Forbid();

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? User.Identity?.Name ?? "system";
            var rows = await LoadTemplateBackfillRowsAsync(cn);

            var result = new TemplateBackfillResultVm
            {
                Total = rows.Count
            };

            foreach (var row in rows)
            {
                var dbPath = FirstNotEmpty(row.FilePath, row.ExcelFilePath);
                var absolutePath = ToSafeAbsolutePath(dbPath);

                if (string.IsNullOrWhiteSpace(absolutePath) || !System.IO.File.Exists(absolutePath))
                {
                    result.FileMissing++;
                    result.Items.Add(new TemplateBackfillItemVm
                    {
                        VersionId = row.VersionId,
                        Status = "FileMissing",
                        Path = dbPath
                    });
                    continue;
                }

                if (!string.Equals(row.Storage, "Disk", StringComparison.OrdinalIgnoreCase))
                {
                    result.Skipped++;
                    result.Items.Add(new TemplateBackfillItemVm
                    {
                        VersionId = row.VersionId,
                        Status = "SkippedStorage",
                        Path = dbPath
                    });
                    continue;
                }

                string? backupPath = null;

                try
                {
                    backupPath = _templateVersionPrepareService.CreateBackupFile(absolutePath);

                    var prepare = _templateVersionPrepareService.PrepareExistingVersionFile(
                        cn,
                        row.VersionId,
                        absolutePath
                    );

                    var preparedAt = DateTime.Now;

                    await using var tx = await cn.BeginTransactionAsync();

                    try
                    {
                        await using (var cmd = cn.CreateCommand())
                        {
                            cmd.Transaction = (SqlTransaction)tx;
                            cmd.CommandText = @"
UPDATE dbo.DocTemplateVersion
   SET ExcelFileSize = @ExcelFileSize,
       PreviewJson = @PreviewJson,
       PreparedAt = @PreparedAt,
       TemplateFileHash = @TemplateFileHash,
       ProtectionRuleCode = @ProtectionRuleCode,
       VisualMetricRuleCode = @VisualMetricRuleCode,
       VisualSource = @VisualSource,
       VisualRangeA1 = @VisualRangeA1,
       VisualWidthPx = @VisualWidthPx,
       VisualHeightPx = @VisualHeightPx
 WHERE Id = @VersionId;";

                            cmd.Parameters.Add("@ExcelFileSize", SqlDbType.BigInt).Value = prepare.ExcelFileSize;
                            cmd.Parameters.Add("@PreviewJson", SqlDbType.NVarChar, -1).Value = prepare.PreviewJson;
                            cmd.Parameters.Add("@PreparedAt", SqlDbType.DateTime2).Value = preparedAt;
                            cmd.Parameters.Add("@TemplateFileHash", SqlDbType.Char, 64).Value = prepare.TemplateFileHash;
                            cmd.Parameters.Add("@ProtectionRuleCode", SqlDbType.NVarChar, 50).Value = prepare.ProtectionRuleCode;
                            cmd.Parameters.Add("@VisualMetricRuleCode", SqlDbType.NVarChar, 50).Value = prepare.VisualMetricRuleCode;
                            cmd.Parameters.Add("@VisualSource", SqlDbType.NVarChar, 50).Value = prepare.VisualSource;
                            cmd.Parameters.Add("@VisualRangeA1", SqlDbType.NVarChar, 100).Value = prepare.VisualRangeA1;
                            cmd.Parameters.Add("@VisualWidthPx", SqlDbType.Int).Value = prepare.VisualWidthPx;
                            cmd.Parameters.Add("@VisualHeightPx", SqlDbType.Int).Value = prepare.VisualHeightPx;
                            cmd.Parameters.Add("@VersionId", SqlDbType.BigInt).Value = row.VersionId;

                            await cmd.ExecuteNonQueryAsync();
                        }

                        await using (var cmd = cn.CreateCommand())
                        {
                            cmd.Transaction = (SqlTransaction)tx;
                            cmd.CommandText = @"
UPDATE dbo.DocTemplateFile
   SET FileSize = @ExcelFileSize,
       FileSizeBytes = @ExcelFileSize
 WHERE VersionId = @VersionId
   AND FileRole = N'ExcelFile';";

                            cmd.Parameters.Add("@ExcelFileSize", SqlDbType.BigInt).Value = prepare.ExcelFileSize;
                            cmd.Parameters.Add("@VersionId", SqlDbType.BigInt).Value = row.VersionId;

                            await cmd.ExecuteNonQueryAsync();
                        }

                        await using (var cmd = cn.CreateCommand())
                        {
                            cmd.Transaction = (SqlTransaction)tx;
                            cmd.CommandText = @"
IF EXISTS (SELECT 1 FROM dbo.DocTemplateFile WHERE VersionId = @VersionId AND FileRole = N'PreviewJson')
BEGIN
    UPDATE dbo.DocTemplateFile
       SET Contents = @PreviewJson,
           FileSize = DATALENGTH(@PreviewJson),
           FileSizeBytes = DATALENGTH(@PreviewJson),
           ContentType = N'application/json'
     WHERE VersionId = @VersionId
       AND FileRole = N'PreviewJson';
END
ELSE
BEGIN
    INSERT dbo.DocTemplateFile
           (TemplateId, VersionId, FileRole, Storage, FileName, FilePath,
            FileSize, FileSizeBytes, ContentType, Contents, CreatedAt, CreatedBy)
    VALUES (@TemplateId, @VersionId, N'PreviewJson', N'Db', N'preview.json', NULL,
            DATALENGTH(@PreviewJson), DATALENGTH(@PreviewJson), N'application/json', @PreviewJson,
            SYSUTCDATETIME(), @CreatedBy);
END";

                            cmd.Parameters.Add("@TemplateId", SqlDbType.Int).Value = row.TemplateId;
                            cmd.Parameters.Add("@VersionId", SqlDbType.BigInt).Value = row.VersionId;
                            cmd.Parameters.Add("@PreviewJson", SqlDbType.NVarChar, -1).Value = prepare.PreviewJson;
                            cmd.Parameters.Add("@CreatedBy", SqlDbType.NVarChar, 100).Value = userId;

                            await cmd.ExecuteNonQueryAsync();
                        }

                        await tx.CommitAsync();
                    }
                    catch
                    {
                        await tx.RollbackAsync();
                        TemplateVersionPrepareService.RestoreBackupFile(backupPath, absolutePath);
                        throw;
                    }

                    result.Success++;
                    result.Items.Add(new TemplateBackfillItemVm
                    {
                        VersionId = row.VersionId,
                        Status = "Success",
                        Path = dbPath,
                        VisualRangeA1 = prepare.VisualRangeA1,
                        VisualWidthPx = prepare.VisualWidthPx,
                        VisualHeightPx = prepare.VisualHeightPx
                    });
                }
                catch
                {
                    if (!string.IsNullOrWhiteSpace(backupPath))
                        TemplateVersionPrepareService.RestoreBackupFile(backupPath, absolutePath);

                    result.Failed++;
                    result.Items.Add(new TemplateBackfillItemVm
                    {
                        VersionId = row.VersionId,
                        Status = "Failed",
                        Path = dbPath
                    });
                }
            }

            return Json(new
            {
                ok = result.Failed == 0,
                result.Total,
                result.Success,
                result.FileMissing,
                result.Skipped,
                result.Failed,
                result.Items
            });
        }

        private async Task<List<DocTemplateManageRowVm>> LoadLatestTemplateRowsAsync(SqlConnection cn)
        {
            const string sql = @"
SELECT
    m.Id AS TemplateId,
    v.Id AS VersionId,
    f.Id AS FileId,
    m.CompCd,
    m.DepartmentId,
    m.KindCode,
    COALESCE(NULLIF(LTRIM(RTRIM(tkm.Name)), N''), NULLIF(LTRIM(RTRIM(m.KindCode)), N'')) AS KindName,
    m.DocCode,
    m.DocName,
    m.Title,
    m.ApprovalCount,
    m.IsActive,
    v.VersionNo,
    v.VisualRangeA1,
    v.VisualWidthPx,
    v.VisualHeightPx,
    v.ExcelFileName,
    v.ExcelStorage,
    v.ExcelFilePath,
    v.ExcelFileSize,
    v.ExcelContentType,
    v.CreatedBy AS VersionCreatedBy,
    v.CreatedAt AS VersionCreatedAt,
    f.Storage AS FileStorage,
    f.FilePath AS FilePath,
    f.FileName AS FileName,
    f.ContentType AS FileContentType,
    f.FileSizeBytes,
    CASE
        WHEN f.Blob IS NULL THEN CONVERT(bigint, 0)
        ELSE CONVERT(bigint, DATALENGTH(f.Blob))
    END AS BlobSize,
    f.CreatedBy AS FileCreatedBy,
    f.CreatedAt AS FileCreatedAt,
    m.CreatedBy AS MasterCreatedBy,
    m.CreatedAt AS MasterCreatedAt,
    m.UpdatedBy AS MasterUpdatedBy,
    m.UpdatedAt AS MasterUpdatedAt,
    COALESCE
    (
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), cm.TimeZoneId))), N''),
        N'Asia/Seoul'
    ) AS CompTimeZoneId,
    COALESCE
    (
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(20), cm.Locale))), N''),
        N''
    ) AS CompLocale,
    ISNULL(vc.VersionCount, 0) AS VersionCount,
    COALESCE(mcup.DisplayName, mcu.DisplayName, m.CreatedBy) AS CreatedByName,
    COALESCE
    (
        NULLIF(LTRIM(RTRIM(muup.DisplayName)), N''),
        NULLIF(LTRIM(RTRIM(muu.UserName)), N''),
        NULLIF(LTRIM(RTRIM(m.UpdatedBy)), N'')
    ) AS UpdatedByName
FROM dbo.DocTemplateMaster m
OUTER APPLY
(
    SELECT TOP 1 *
    FROM dbo.DocTemplateVersion x
    WHERE x.TemplateId = m.Id
    ORDER BY ISNULL(x.VersionNo, 0) DESC, x.Id DESC
) v
OUTER APPLY
(
    SELECT COUNT(1) AS VersionCount
    FROM dbo.DocTemplateVersion x
    WHERE x.TemplateId = m.Id
) vc
OUTER APPLY
(
    SELECT TOP 1 *
    FROM dbo.DocTemplateFile x
    WHERE x.TemplateId = m.Id
      AND x.VersionId = v.Id
      AND x.FileRole = N'ExcelFile'
    ORDER BY x.Id DESC
) f
LEFT JOIN dbo.CompMasters cm
       ON cm.CompCd = m.CompCd
LEFT JOIN dbo.TemplateKindMasters tkm
       ON tkm.CompCd = m.CompCd
      AND tkm.DepartmentId = m.DepartmentId
      AND tkm.Code = m.KindCode
LEFT JOIN dbo.AspNetUsers mcu
       ON mcu.Id = m.CreatedBy
LEFT JOIN dbo.UserProfiles mcup
       ON mcup.UserId = m.CreatedBy
OUTER APPLY
(
    SELECT TOP 1 u.*
    FROM dbo.AspNetUsers u
    WHERE NULLIF(LTRIM(RTRIM(m.UpdatedBy)), N'') IS NOT NULL
      AND
      (
             u.Id = m.UpdatedBy
          OR u.UserName = m.UpdatedBy
          OR u.NormalizedUserName = UPPER(m.UpdatedBy)
          OR u.Email = m.UpdatedBy
          OR u.NormalizedEmail = UPPER(m.UpdatedBy)
      )
    ORDER BY
        CASE
            WHEN u.Id = m.UpdatedBy THEN 0
            WHEN u.UserName = m.UpdatedBy THEN 1
            WHEN u.NormalizedUserName = UPPER(m.UpdatedBy) THEN 2
            WHEN u.Email = m.UpdatedBy THEN 3
            WHEN u.NormalizedEmail = UPPER(m.UpdatedBy) THEN 4
            ELSE 9
        END
) muu
LEFT JOIN dbo.UserProfiles muup
       ON muup.UserId = muu.Id
ORDER BY
    m.CompCd ASC,
    m.DepartmentId ASC,
    tkm.SortOrder ASC,
    m.KindCode ASC,
    m.DocName ASC,
    m.DocCode ASC;";

            var list = new List<DocTemplateManageRowVm>();

            await using var cmd = cn.CreateCommand();
            cmd.CommandText = sql;

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var dbPath = FirstNotEmpty(
                    ReadString(rd, "FilePath"),
                    ReadString(rd, "ExcelFilePath")
                );

                var fileName = FirstNotEmpty(
                    ReadString(rd, "FileName"),
                    ReadString(rd, "ExcelFileName")
                );

                var storage = FirstNotEmpty(
                    ReadString(rd, "FileStorage"),
                    ReadString(rd, "ExcelStorage")
                );

                var absolutePath = ToSafeAbsolutePath(dbPath);
                var diskFileExists = !string.IsNullOrWhiteSpace(absolutePath) && System.IO.File.Exists(absolutePath);
                var blobSize = ReadLong(rd, "BlobSize") ?? 0;
                var isDbStorage = string.Equals(storage, "Db", StringComparison.OrdinalIgnoreCase);
                var fileExists = isDbStorage ? blobSize > 0 : diskFileExists;
                var isUnderTemplateRoot = IsUnderDocTemplatesRoot(dbPath, absolutePath);

                var compTimeZoneId = ReadString(rd, "CompTimeZoneId");
                var compLocale = ReadString(rd, "CompLocale");

                var createdAtUtc = FirstNotNull(
                    ReadDateTime(rd, "FileCreatedAt"),
                    ReadDateTime(rd, "VersionCreatedAt"),
                    ReadDateTime(rd, "MasterCreatedAt")
                );

                var updatedAtUtc = ReadDateTime(rd, "MasterUpdatedAt");

                var createdAtLocal = DocControllerHelper.ConvertUtcToLocal(createdAtUtc, compTimeZoneId);
                var updatedAtLocal = DocControllerHelper.ConvertUtcToLocal(updatedAtUtc, compTimeZoneId);

                var row = new DocTemplateManageRowVm
                {
                    TemplateId = ReadInt(rd, "TemplateId") ?? 0,
                    VersionId = ReadLong(rd, "VersionId"),
                    FileId = ReadLong(rd, "FileId"),
                    CompCd = ReadString(rd, "CompCd"),
                    DepartmentId = ReadInt(rd, "DepartmentId"),
                    KindCode = ReadString(rd, "KindCode"),
                    KindName = ReadString(rd, "KindName"),
                    DocCode = ReadString(rd, "DocCode"),
                    DocName = ReadString(rd, "DocName"),
                    Title = ReadString(rd, "Title"),
                    ApprovalCount = ReadInt(rd, "ApprovalCount") ?? 0,
                    IsActive = ReadBool(rd, "IsActive"),
                    VersionNo = ReadInt(rd, "VersionNo"),
                    VisualRangeA1 = ReadString(rd, "VisualRangeA1"),
                    VisualWidthPx = ReadInt(rd, "VisualWidthPx"),
                    VisualHeightPx = ReadInt(rd, "VisualHeightPx"),
                    VersionCount = ReadInt(rd, "VersionCount") ?? 0,
                    FileName = fileName,

                    Storage = storage,
                    ContentType = FirstNotEmpty(ReadString(rd, "FileContentType"), ReadString(rd, "ExcelContentType")),
                    FileSizeBytes = FirstNotNull(
                        ReadLong(rd, "FileSizeBytes"),
                        ReadLong(rd, "ExcelFileSize"),
                        blobSize > 0 ? blobSize : null
                    ),
                    CurrentPath = NormalizeDbPath(dbPath),
                    AbsolutePath = absolutePath,
                    FileExists = fileExists,
                    IsUnderTemplateRoot = isUnderTemplateRoot,
                    PathStatusKey = GetPathStatusKey(storage, dbPath, absolutePath, fileExists, isUnderTemplateRoot),
                    CreatedBy = FirstNotEmpty(ReadString(rd, "FileCreatedBy"), ReadString(rd, "VersionCreatedBy"), ReadString(rd, "MasterCreatedBy")),
                    CreatedByName = ReadString(rd, "CreatedByName"),
                    CreatedAtUtc = DocControllerHelper.TreatAsUtc(createdAtUtc),
                    CreatedAt = createdAtLocal,
                    CreatedAtLocalText = DocControllerHelper.FormatLocalMinute(createdAtLocal),
                    CreatedAtLocalDateKey = DocControllerHelper.FormatLocalDateKey(createdAtLocal),
                    UpdatedBy = ReadString(rd, "MasterUpdatedBy"),
                    UpdatedByName = ReadString(rd, "UpdatedByName"),
                    UpdatedAtUtc = DocControllerHelper.TreatAsUtc(updatedAtUtc),
                    UpdatedAt = updatedAtLocal,
                    UpdatedAtLocalText = DocControllerHelper.FormatLocalMinute(updatedAtLocal),
                    UpdatedAtLocalDateKey = DocControllerHelper.FormatLocalDateKey(updatedAtLocal),
                    CompTimeZoneId = compTimeZoneId,
                    CompLocale = compLocale,
                    FileLastWriteAt = diskFileExists ? System.IO.File.GetLastWriteTime(absolutePath!) : null
                };

                list.Add(row);
            }

            await ApplyDisplayNamesAsync(cn, list);

            return list;
        }

        private async Task ApplyDisplayNamesAsync(SqlConnection cn, List<DocTemplateManageRowVm> list)
        {
            if (list.Count == 0)
                return;

            var compMap = await LoadLookupMapAsync(
                cn,
                new[]
                {
                    new LookupTableSpec("dbo.CompMaster",
                        new[] { "CompCd", "Code", "Cd" },
                        new[] { "Name", "CompName", "CompanyName", "ShortName", "DisplayName" }),
                    new LookupTableSpec("dbo.CompMasters",
                        new[] { "CompCd", "Code", "Cd" },
                        new[] { "Name", "CompName", "CompanyName", "ShortName", "DisplayName" }),
                    new LookupTableSpec("dbo.CompanyMaster",
                        new[] { "CompCd", "Code", "Cd" },
                        new[] { "Name", "CompName", "CompanyName", "ShortName", "DisplayName" }),
                    new LookupTableSpec("dbo.CompanyMasters",
                        new[] { "CompCd", "Code", "Cd" },
                        new[] { "Name", "CompName", "CompanyName", "ShortName", "DisplayName" })
                });

            var departmentMap = await LoadDepartmentNameMapAsync(cn);

            foreach (var row in list)
            {
                row.CompName = LookupDisplayName(compMap, row.CompCd);

                if (row.DepartmentId.HasValue)
                {
                    if (row.DepartmentId.Value == 0)
                    {
                        row.DepartmentName = "공용";
                    }
                    else
                    {
                        var deptKey = MakeDepartmentMapKey(row.CompCd, row.DepartmentId.Value);
                        row.DepartmentName = departmentMap.TryGetValue(deptKey, out var departmentName)
                            ? departmentName
                            : row.DepartmentId.Value.ToString(CultureInfo.InvariantCulture);
                    }
                }
                else
                {
                    row.DepartmentName = null;
                }

                row.KindName = FirstNotEmpty(row.KindName, row.KindCode);
            }
        }

        private async Task<Dictionary<string, string>> LoadDepartmentNameMapAsync(SqlConnection cn)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (!await TableExistsAsync(cn, "dbo", "DepartmentMasters"))
                return map;

            var hasLoc = await TableExistsAsync(cn, "dbo", "DepartmentMasterLoc");
            var langCode = CultureInfo.CurrentUICulture.Name;
            var twoLetter = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;

            var sql = hasLoc
                ? @"
SELECT
    dm.CompCd,
    dm.Id AS DepartmentId,
    COALESCE
    (
        NULLIF(LTRIM(RTRIM(dml.Name)), N''),
        NULLIF(LTRIM(RTRIM(dml.ShortName)), N''),
        NULLIF(LTRIM(RTRIM(dm.Code)), N''),
        CONVERT(nvarchar(20), dm.Id)
    ) AS DepartmentName
FROM dbo.DepartmentMasters dm
OUTER APPLY
(
    SELECT TOP 1 l.Name, l.ShortName
    FROM dbo.DepartmentMasterLoc l
    WHERE l.DepartmentId = dm.Id
    ORDER BY
        CASE
            WHEN l.LangCode = @LangCode THEN 0
            WHEN l.LangCode = @TwoLetter THEN 1
            WHEN l.LangCode IN (N'ko-KR', N'ko') THEN 2
            ELSE 3
        END
) dml;"
                : @"
SELECT
    dm.CompCd,
    dm.Id AS DepartmentId,
    COALESCE
    (
        NULLIF(LTRIM(RTRIM(dm.Code)), N''),
        CONVERT(nvarchar(20), dm.Id)
    ) AS DepartmentName
FROM dbo.DepartmentMasters dm;";

            await using var cmd = cn.CreateCommand();
            cmd.CommandText = sql;

            if (hasLoc)
            {
                cmd.Parameters.Add("@LangCode", SqlDbType.NVarChar, 20).Value = string.IsNullOrWhiteSpace(langCode) ? "ko-KR" : langCode;
                cmd.Parameters.Add("@TwoLetter", SqlDbType.NVarChar, 10).Value = string.IsNullOrWhiteSpace(twoLetter) ? "ko" : twoLetter;
            }

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var compCd = ReadString(rd, "CompCd");
                var departmentId = ReadInt(rd, "DepartmentId");
                var name = ReadString(rd, "DepartmentName");

                if (!departmentId.HasValue || string.IsNullOrWhiteSpace(name))
                    continue;

                var key = MakeDepartmentMapKey(compCd, departmentId.Value);
                if (!map.ContainsKey(key))
                    map.Add(key, name);
            }

            return map;
        }

        private async Task<Dictionary<string, string>> LoadLookupMapAsync(SqlConnection cn, LookupTableSpec[] specs)
        {
            foreach (var spec in specs)
            {
                var objectName = SplitObjectName(spec.TableName);
                if (!await TableExistsAsync(cn, objectName.Schema, objectName.Table))
                    continue;

                var keyColumn = await FindFirstExistingColumnAsync(cn, objectName.Schema, objectName.Table, spec.KeyColumns);
                var valueColumn = await FindFirstExistingColumnAsync(cn, objectName.Schema, objectName.Table, spec.ValueColumns);

                if (string.IsNullOrWhiteSpace(keyColumn) || string.IsNullOrWhiteSpace(valueColumn))
                    continue;

                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                var sql = $@"
SELECT
    CONVERT(nvarchar(100), {QuoteSqlName(keyColumn)}) AS LookupKey,
    CONVERT(nvarchar(300), {QuoteSqlName(valueColumn)}) AS LookupValue
FROM {QuoteSqlName(objectName.Schema)}.{QuoteSqlName(objectName.Table)}
WHERE {QuoteSqlName(keyColumn)} IS NOT NULL
  AND {QuoteSqlName(valueColumn)} IS NOT NULL;";

                await using var cmd = cn.CreateCommand();
                cmd.CommandText = sql;

                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    var key = ReadString(rd, "LookupKey");
                    var value = ReadString(rd, "LookupValue");

                    if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
                        continue;

                    key = key.Trim();
                    value = value.Trim();

                    if (!map.ContainsKey(key))
                        map.Add(key, value);
                }

                if (map.Count > 0)
                    return map;
            }

            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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
            return Convert.ToInt32(result) == 1;
        }

        private async Task<string?> FindFirstExistingColumnAsync(SqlConnection cn, string schema, string table, IEnumerable<string> candidates)
        {
            foreach (var candidate in candidates)
            {
                await using var cmd = cn.CreateCommand();
                cmd.CommandText = @"
SELECT CASE
       WHEN EXISTS
       (
           SELECT 1
           FROM INFORMATION_SCHEMA.COLUMNS
           WHERE TABLE_SCHEMA = @Schema
             AND TABLE_NAME = @Table
             AND COLUMN_NAME = @Column
       )
       THEN 1 ELSE 0 END;";

                cmd.Parameters.Add("@Schema", SqlDbType.NVarChar, 128).Value = schema;
                cmd.Parameters.Add("@Table", SqlDbType.NVarChar, 128).Value = table;
                cmd.Parameters.Add("@Column", SqlDbType.NVarChar, 128).Value = candidate;

                var result = await cmd.ExecuteScalarAsync();
                if (Convert.ToInt32(result) == 1)
                    return candidate;
            }

            return null;
        }

        private async Task<DownloadFileVm?> LoadDownloadFileAsync(SqlConnection cn, long versionId)
        {
            await using var cmd = cn.CreateCommand();
            cmd.CommandText = @"
SELECT TOP 1
    COALESCE(f.Storage, v.ExcelStorage, N'Disk') AS Storage,
    COALESCE(f.FilePath, v.ExcelFilePath) AS FilePath,
    COALESCE(f.FileName, v.ExcelFileName) AS FileName,
    COALESCE(f.ContentType, v.ExcelContentType) AS ContentType,
    f.Blob
FROM dbo.DocTemplateVersion v
LEFT JOIN dbo.DocTemplateFile f
       ON f.VersionId = v.Id
      AND f.TemplateId = v.TemplateId
      AND f.FileRole = N'ExcelFile'
WHERE v.Id = @VersionId
ORDER BY f.Id DESC;";

            cmd.Parameters.Add("@VersionId", SqlDbType.BigInt).Value = versionId;

            await using var rd = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess);
            if (!await rd.ReadAsync())
                return null;

            byte[]? blob = null;
            var blobOrdinal = rd.GetOrdinal("Blob");
            if (!await rd.IsDBNullAsync(blobOrdinal))
            {
                using var ms = new MemoryStream();
                using var stream = rd.GetStream(blobOrdinal);
                await stream.CopyToAsync(ms);
                blob = ms.ToArray();
            }

            return new DownloadFileVm
            {
                Storage = ReadString(rd, "Storage"),
                FilePath = ReadString(rd, "FilePath"),
                FileName = ReadString(rd, "FileName"),
                ContentType = ReadString(rd, "ContentType"),
                Blob = blob
            };
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
            return Convert.ToInt32(result) == 1;
        }

        private string GetConnectionString()
        {
            var cs = _configuration.GetConnectionString("DefaultConnection");
            if (string.IsNullOrWhiteSpace(cs))
                throw new InvalidOperationException("DefaultConnection is not configured.");

            return cs;
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

        // 2026.06.15 Added: 메타 누락 템플릿 버전 전체를 조회하여 최신 버전 그리드 표시 여부와 무관하게 보정 대상으로 사용한다.
        private async Task<List<TemplateBackfillVersionRow>> LoadTemplateBackfillRowsAsync(SqlConnection cn)
        {
            var list = new List<TemplateBackfillVersionRow>();

            await using var cmd = cn.CreateCommand();
            cmd.CommandText = @"
SELECT
    v.Id AS VersionId,
    v.TemplateId,
    v.VersionNo,
    COALESCE(f.Storage, v.ExcelStorage, N'Disk') AS Storage,
    COALESCE(f.FilePath, v.ExcelFilePath) AS FilePath,
    v.ExcelFilePath
FROM dbo.DocTemplateVersion v
OUTER APPLY
(
    SELECT TOP 1
        x.Storage,
        x.FilePath
    FROM dbo.DocTemplateFile x
    WHERE x.TemplateId = v.TemplateId
      AND x.VersionId = v.Id
      AND x.FileRole = N'ExcelFile'
    ORDER BY x.Id DESC
) f
WHERE v.PreparedAt IS NULL
   OR v.TemplateFileHash IS NULL
   OR v.ProtectionRuleCode IS NULL
   OR v.VisualMetricRuleCode IS NULL
   OR v.VisualSource IS NULL
   OR v.VisualRangeA1 IS NULL
   OR v.VisualWidthPx IS NULL
   OR v.VisualHeightPx IS NULL
ORDER BY v.Id ASC;";

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                list.Add(new TemplateBackfillVersionRow
                {
                    VersionId = ReadLong(rd, "VersionId") ?? 0,
                    TemplateId = ReadInt(rd, "TemplateId") ?? 0,
                    VersionNo = ReadInt(rd, "VersionNo") ?? 0,
                    Storage = ReadString(rd, "Storage"),
                    FilePath = ReadString(rd, "FilePath"),
                    ExcelFilePath = ReadString(rd, "ExcelFilePath")
                });
            }

            return list.Where(x => x.VersionId > 0 && x.TemplateId > 0).ToList();
        }

        private bool IsUnderDocTemplatesRoot(string? dbPath, string? absolutePath)
        {
            var relative = NormalizeDbPath(dbPath);
            if (!string.IsNullOrWhiteSpace(relative))
            {
                var normalizedRelative = relative.Replace('/', '\\').TrimStart('\\');
                if (normalizedRelative.StartsWith(@"App_Data\DocTemplates\", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            if (string.IsNullOrWhiteSpace(absolutePath))
                return false;

            var docTemplatesRoot = Path.GetFullPath(Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates"));
            var full = Path.GetFullPath(absolutePath);

            return full.StartsWith(docTemplatesRoot, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetPathStatusKey(string? storage, string? dbPath, string? absolutePath, bool fileExists, bool isUnderTemplateRoot)
        {
            if (string.Equals(storage, "Db", StringComparison.OrdinalIgnoreCase))
                return fileExists ? "DTM_PathStatus_Exists" : "DTM_Option_FileMissing";

            if (string.IsNullOrWhiteSpace(dbPath))
                return "DTM_PathStatus_NoPath";

            if (string.IsNullOrWhiteSpace(absolutePath))
                return "DTM_PathStatus_UnsafePath";

            if (!isUnderTemplateRoot)
                return "DTM_PathStatus_OutOfRoot";

            if (!fileExists)
                return "DTM_Option_FileMissing";

            return "DTM_PathStatus_Exists";
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

        private static string? LookupDisplayName(Dictionary<string, string> map, string? code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return null;

            var key = code.Trim();
            return map.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
                ? value
                : key;
        }

        private static string MakeDepartmentMapKey(string? compCd, int departmentId)
        {
            return (compCd ?? "").Trim() + "|" + departmentId.ToString(CultureInfo.InvariantCulture);
        }

        private static (string Schema, string Table) SplitObjectName(string objectName)
        {
            var parts = objectName.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            return parts.Length == 2
                ? (parts[0], parts[1])
                : ("dbo", objectName);
        }

        private static string QuoteSqlName(string value)
        {
            return "[" + value.Replace("]", "]]") + "]";
        }

        private static string? FirstNotEmpty(params string?[] values)
        {
            return values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
        }

        private static long? FirstNotNull(params long?[] values)
        {
            return values.FirstOrDefault(x => x.HasValue);
        }

        private static DateTime? FirstNotNull(params DateTime?[] values)
        {
            return values.FirstOrDefault(x => x.HasValue);
        }

        private static string? ReadString(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            return rd.IsDBNull(ordinal) ? null : rd.GetString(ordinal);
        }

        private static int? ReadInt(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            return rd.IsDBNull(ordinal) ? null : rd.GetInt32(ordinal);
        }

        private static long? ReadLong(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            return rd.IsDBNull(ordinal) ? null : rd.GetInt64(ordinal);
        }

        private static bool ReadBool(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            return !rd.IsDBNull(ordinal) && rd.GetBoolean(ordinal);
        }

        private static DateTime? ReadDateTime(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            return rd.IsDBNull(ordinal) ? null : rd.GetDateTime(ordinal);
        }

        public sealed class DocTemplateManageIndexVm
        {
            public DocTemplateManageSearchVm Search { get; set; } = new();
            public List<DocTemplateManageRowVm> Items { get; set; } = new();
        }

        public sealed class DocTemplateManageSearchVm
        {
            public string? CompCd { get; set; }
            public int? DepartmentId { get; set; }
            public string? KindCode { get; set; }
            public string? UseState { get; set; }
            public string? PathState { get; set; }
            public string? Keyword { get; set; }
        }

        public sealed class DocTemplateManageRowVm
        {
            public int TemplateId { get; set; }
            public long? VersionId { get; set; }
            public long? FileId { get; set; }
            public string? CompCd { get; set; }
            public string? CompName { get; set; }
            public string? CompTimeZoneId { get; set; }
            public string? CompLocale { get; set; }
            public int? DepartmentId { get; set; }
            public string? DepartmentName { get; set; }
            public string? KindCode { get; set; }
            public string? KindName { get; set; }
            public string? DocCode { get; set; }
            public string? DocName { get; set; }
            public string? Title { get; set; }
            public int ApprovalCount { get; set; }
            public bool IsActive { get; set; }
            public int? VersionNo { get; set; }
            public string? VisualRangeA1 { get; set; }
            public int? VisualWidthPx { get; set; }
            public int? VisualHeightPx { get; set; }
            public int VersionCount { get; set; }
            public string? FileName { get; set; }
            public string? Storage { get; set; }
            public string? ContentType { get; set; }
            public long? FileSizeBytes { get; set; }
            public string? CurrentPath { get; set; }
            public string? AbsolutePath { get; set; }
            public bool FileExists { get; set; }
            public bool IsUnderTemplateRoot { get; set; }
            public string PathStatusKey { get; set; } = "DTM_PathStatus_NoPath";
            public string? CreatedBy { get; set; }
            public string? CreatedByName { get; set; }
            public DateTime? CreatedAtUtc { get; set; }
            public DateTime? CreatedAt { get; set; }
            public string? CreatedAtLocalText { get; set; }
            public string? CreatedAtLocalDateKey { get; set; }
            public string? UpdatedBy { get; set; }
            public string? UpdatedByName { get; set; }
            public DateTime? UpdatedAtUtc { get; set; }
            public DateTime? UpdatedAt { get; set; }
            public string? UpdatedAtLocalText { get; set; }
            public string? UpdatedAtLocalDateKey { get; set; }
            public DateTime? FileLastWriteAt { get; set; }
        }

        public sealed class SaveUseStateRequest
        {
            public List<SaveUseStateItem> Items { get; set; } = new();
        }

        public sealed class SaveUseStateItem
        {
            public int TemplateId { get; set; }
            public bool IsActive { get; set; }
        }

        private sealed class DownloadFileVm
        {
            public string? Storage { get; set; }
            public string? FilePath { get; set; }
            public string? FileName { get; set; }
            public string? ContentType { get; set; }
            public byte[]? Blob { get; set; }
        }

        // 2026.06.15 Added: 템플릿 버전 일괄 보정 대상 행 정보를 보관한다.
        private sealed class TemplateBackfillVersionRow
        {
            public long VersionId { get; set; }
            public int TemplateId { get; set; }
            public int VersionNo { get; set; }
            public string? Storage { get; set; }
            public string? FilePath { get; set; }
            public string? ExcelFilePath { get; set; }
        }

        // 2026.06.15 Added: 템플릿 버전 일괄 보정 결과를 반환한다.
        private sealed class TemplateBackfillResultVm
        {
            public int Total { get; set; }
            public int Success { get; set; }
            public int FileMissing { get; set; }
            public int Skipped { get; set; }
            public int Failed { get; set; }
            public List<TemplateBackfillItemVm> Items { get; set; } = new();
        }

        // 2026.06.15 Added: 템플릿 버전별 일괄 보정 처리 결과를 보관한다.
        private sealed class TemplateBackfillItemVm
        {
            public long VersionId { get; set; }
            public string Status { get; set; } = string.Empty;
            public string? Path { get; set; }
            public string? VisualRangeA1 { get; set; }
            public int VisualWidthPx { get; set; }
            public int VisualHeightPx { get; set; }
        }

        private sealed class LookupTableSpec
        {
            public LookupTableSpec(string tableName, string[] keyColumns, string[] valueColumns)
            {
                TableName = tableName;
                KeyColumns = keyColumns;
                ValueColumns = valueColumns;
            }

            public string TableName { get; }
            public string[] KeyColumns { get; }
            public string[] ValueColumns { get; }
        }
    }
}