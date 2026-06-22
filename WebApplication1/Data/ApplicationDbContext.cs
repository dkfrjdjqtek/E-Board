using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using WebApplication1.Models;
using WebApplication1.Data.Config;
using WebApplication1.Controllers;

namespace WebApplication1.Data;

public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
{
    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options) { }

    public DbSet<WebAuthnCredential> WebAuthnCredentials => Set<WebAuthnCredential>();
    public DbSet<UserProfile> UserProfiles => Set<UserProfile>();
    public DbSet<DepartmentMaster> DepartmentMasters => Set<DepartmentMaster>();
    public DbSet<PositionMaster> PositionMasters => Set<PositionMaster>();
    public DbSet<DepartmentMasterLoc> DepartmentMasterLoc => Set<DepartmentMasterLoc>();
    public DbSet<PositionMasterLoc> PositionMasterLoc => Set<PositionMasterLoc>();
    public DbSet<CompMaster> CompMasters => Set<CompMaster>();

    //  새 테이블 DbSet
    public DbSet<TemplateKindMaster> TemplateKindMasters { get; set; } = null!;
    public DbSet<TemplateKindMasterLoc> TemplateKindMasterLoc { get; set; } = null!;
    // 2025.09.25 Added: 초대 메일 발송 이력 테이블 DbSet 등록 PK만 사용 FK 미생성
    public DbSet<InviteAudit> InviteAudits { get; set; } = default!;

    public DbSet<DocTemplateMaster> DocTemplateMasters { get; set; } = default!;
    public DbSet<DocTemplateVersion> DocTemplateVersions { get; set; } = default!;
    public DbSet<DocTemplateFile> DocTemplateFiles { get; set; } = default!;
    public DbSet<DocTemplateApproval> DocTemplateApprovals { get; set; } = default!;

    // 2026.06.16 Added: 전결 테이블 DbSet 등록 Contents 전결 규칙과 금액 조건 및 문서별 전결 적용 결과를 조회 저장
    public DbSet<DocTemplateDelegationRule> DocTemplateDelegationRules { get; set; } = default!;
    public DbSet<DocTemplateDelegationAmountRule> DocTemplateDelegationAmountRules { get; set; } = default!;
    public DbSet<DocumentDelegationResult> DocumentDelegationResults { get; set; } = default!;

    public DbSet<WebPushSubscription> WebPushSubscriptions { get; set; } = default!;

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // ★ 컨버터: Required 필드에만 사용 (null 가정 없음)
        var UpperTrim = new ValueConverter<string, string>(
            v => v.Trim().ToUpperInvariant(),
            v => v);
        var LowerTrim = new ValueConverter<string, string>(
            v => v.Trim().ToLowerInvariant(),
            v => v);

        modelBuilder.HasDefaultSchema("dbo");

        modelBuilder.Entity<PositionMaster>().ToTable("PositionMasters", "dbo");
        modelBuilder.Entity<DepartmentMaster>().ToTable("DepartmentMasters", "dbo");
        modelBuilder.Entity<UserProfile>().ToTable("UserProfiles", "dbo");
        modelBuilder.Entity<PositionMasterLoc>().ToTable("PositionMasterLoc", "dbo");
        modelBuilder.Entity<DepartmentMasterLoc>().ToTable("DepartmentMasterLoc", "dbo");
        modelBuilder.Entity<WebAuthnCredential>().ToTable("WebAuthnCredentials", "dbo");

        // 기존 개별 Config 사용 시 충돌 가능하므로 주석 처리
        // modelBuilder.ApplyConfiguration(new TemplateKindMasterConfig());
        // modelBuilder.ApplyConfiguration(new TemplateKindMasterLocConfig());
        // 또는 프로젝트에서 이미 사용 중이라면
        // modelBuilder.ApplyConfigurationsFromAssembly(typeof(ApplicationDbContext).Assembly);

        // ========== WebAuthnCredential ==========
        modelBuilder.Entity<WebAuthnCredential>(b =>
        {
            b.ToTable("WebAuthnCredentials");
            b.HasKey(x => x.Id);

            b.Property(x => x.UserId).IsRequired().HasMaxLength(450);
            b.Property(x => x.CredentialId).IsRequired();
            b.Property(x => x.CredentialIdHash).IsRequired();

            b.HasIndex(x => x.CredentialIdHash).IsUnique();
            b.HasIndex(x => x.UserId);
            b.HasIndex(x => new { x.UserId, x.Nickname }).IsUnique()
             .HasFilter("[Nickname] IS NOT NULL");

            b.Property(x => x.PublicKey).IsRequired();
            b.Property(x => x.CredType).HasMaxLength(20).HasDefaultValue("public-key");
            b.Property(x => x.Transports).HasMaxLength(200);
            b.Property(x => x.Nickname).HasMaxLength(100);
            b.Property(x => x.CreatedAtUtc).HasDefaultValueSql("SYSUTCDATETIME()");

            b.HasOne<ApplicationUser>()
             .WithMany()
             .HasForeignKey(x => x.UserId)
             .OnDelete(DeleteBehavior.NoAction);
        });

        // ========== UserProfile(1:1) ==========
        modelBuilder.Entity<UserProfile>(e =>
        {
            e.HasKey(p => p.UserId);

            e.HasIndex(x => x.UserId).IsUnique();
            e.Property(x => x.CompCd).HasMaxLength(10).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.DisplayName).HasMaxLength(64);

            e.HasOne(p => p.User)
             .WithOne()
             .HasForeignKey<UserProfile>(p => p.UserId)
             .OnDelete(DeleteBehavior.Cascade);

            e.HasOne(p => p.Department).WithMany()
             .HasForeignKey(p => p.DepartmentId)
             .OnDelete(DeleteBehavior.SetNull);

            e.HasOne(p => p.Position).WithMany()
             .HasForeignKey(p => p.PositionId)
             .OnDelete(DeleteBehavior.SetNull);

            e.Property(x => x.IsAdmin).HasDefaultValue(0);
            e.Property<byte[]>("RowVersion").IsRowVersion();
        });

        modelBuilder.Entity<ApplicationUser>()
            .Property(u => u.IsAdmin).HasDefaultValue(0);

        // --- CompMaster ---
        modelBuilder.Entity<CompMaster>(e =>
        {
            e.HasKey(x => x.CompCd);
            e.Property(x => x.CompCd).HasMaxLength(10);
            e.Property(x => x.Name).HasMaxLength(100).IsRequired();
        });

        // ========== DepartmentMaster ==========
        modelBuilder.Entity<DepartmentMaster>(e =>
        {
            e.Property(x => x.CompCd).HasMaxLength(10).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.Code).HasMaxLength(32).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.Name).HasMaxLength(64).IsRequired();
            e.Property(x => x.IsActive).HasDefaultValue(true);
            e.Property(x => x.SortOrder).HasDefaultValue(0);

            e.HasIndex(x => new { x.CompCd, x.Code }).IsUnique();
            e.HasIndex(x => new { x.CompCd, x.IsActive, x.SortOrder });

            e.ToTable(tb =>
            {
                tb.HasCheckConstraint("CK_DepartmentMasters_Code", "LEN(LTRIM(RTRIM([Code]))) > 0");
            });

            e.Property<byte[]>("RowVersion").IsRowVersion();
        });

        // ========== PositionMaster ==========
        modelBuilder.Entity<PositionMaster>(e =>
        {
            e.Property(x => x.CompCd).HasMaxLength(10).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.Code).HasMaxLength(32).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.Name).HasMaxLength(64).IsRequired();

            e.Property(x => x.RankLevel).HasDefaultValue((short)0);
            e.Property(x => x.IsApprover).HasDefaultValue(false);
            e.Property(x => x.SortOrder).HasDefaultValue(0);

            e.HasIndex(x => new { x.CompCd, x.Code }).IsUnique();
            e.HasIndex(x => new { x.CompCd, x.IsActive, x.RankLevel, x.SortOrder });

            e.ToTable(tb =>
            {
                tb.HasCheckConstraint("CK_PositionMasters_RankLevel", "[RankLevel] >= 0");
            });

            e.Property<byte[]>("RowVersion").IsRowVersion();
        });

        // ========== DepartmentMasterLoc (i18n) ==========
        modelBuilder.Entity<DepartmentMasterLoc>(e =>
        {
            e.Property(x => x.LangCode).HasMaxLength(10).IsRequired().HasConversion(LowerTrim);
            e.Property(x => x.Name).HasMaxLength(64).IsRequired();
            e.Property(x => x.ShortName).HasMaxLength(32);

            e.HasIndex(x => new { x.DepartmentId, x.LangCode }).IsUnique();
            e.HasIndex(x => x.LangCode);

            e.HasOne(x => x.Department)
             .WithMany(m => m.Locs)
             .HasForeignKey(x => x.DepartmentId)
             .OnDelete(DeleteBehavior.Cascade);

            e.Property<byte[]>("RowVersion").IsRowVersion();
        });

        // ========== PositionMasterLoc (i18n) ==========
        modelBuilder.Entity<PositionMasterLoc>(e =>
        {
            e.Property(x => x.LangCode).HasMaxLength(10).IsRequired().HasConversion(LowerTrim);
            e.Property(x => x.Name).HasMaxLength(64).IsRequired();
            e.Property(x => x.ShortName).HasMaxLength(32);

            e.HasIndex(x => new { x.PositionId, x.LangCode }).IsUnique();
            e.HasIndex(x => x.LangCode);

            e.HasOne(x => x.Position)
             .WithMany(m => m.Locs)
             .HasForeignKey(x => x.PositionId)
             .OnDelete(DeleteBehavior.Cascade);

            e.Property<byte[]>("RowVersion").IsRowVersion();
        });

        // ============================================================
        // ========== TemplateKindMaster / TemplateKindMasterLoc ======
        // DB 스크립트에 맞춘 정확한 매핑 (Loc에 CompCd/DepartmentId/RowVersion 포함)
        // ============================================================

        // TemplateKindMasters
        modelBuilder.Entity<TemplateKindMaster>(e =>
        {
            e.ToTable("TemplateKindMasters", "dbo");
            e.HasKey(x => x.Id);

            e.Property(x => x.CompCd).HasMaxLength(10).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.DepartmentId).IsRequired().HasDefaultValue(0);
            e.Property(x => x.Code).HasMaxLength(32).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.Name).HasMaxLength(64).IsRequired();
            e.Property(x => x.IsActive).HasDefaultValue(true);
            e.Property(x => x.SortOrder).HasDefaultValue(0);

            e.HasIndex(x => new { x.CompCd, x.Code }).IsUnique().HasDatabaseName("UX_TemplateKindMasters_CompCd_Code");
            e.HasIndex(x => new { x.CompCd, x.DepartmentId }).HasDatabaseName("IX_TemplateKindMasters_CompCd_Department");

            // RowVersion: 섀도우 속성
            e.Property<byte[]>("RowVersion").IsRowVersion();
        });

        // --- TemplateKindMasterLoc ---
        modelBuilder.Entity<TemplateKindMasterLoc>(e =>
        {
            e.ToTable("TemplateKindMasterLoc", "dbo"); // 트리거 제거 시 OUTPUT 설정 필요 없음

            // PK = (Id, LangCode)
            e.HasKey(x => new { x.Id, x.LangCode });

            // FK(Id) → Masters(Id)  (TemplateKindMasterId 같은 컬럼/속성 사용 금지)
            e.HasOne<TemplateKindMaster>()
             .WithMany()
             .HasForeignKey(x => x.Id)
             .OnDelete(DeleteBehavior.Cascade);

            e.Property(x => x.CompCd).HasMaxLength(10).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.DepartmentId).IsRequired().HasDefaultValue(0);
            e.Property(x => x.LangCode).HasMaxLength(10).IsRequired().HasConversion(LowerTrim);
            e.Property(x => x.Name).HasMaxLength(64).IsRequired();

            // 보조 인덱스 (검색/조인용)
            e.HasIndex(x => new { x.CompCd, x.DepartmentId, x.LangCode });

            // RowVersion: 섀도우 속성
            e.Property<byte[]>("RowVersion").IsRowVersion();
        });

        modelBuilder.Entity<DocTemplateMaster>(e =>
        {
            e.ToTable("DocTemplateMaster", "dbo");
            e.HasKey(x => x.Id);
            e.Property(x => x.CompCd).HasMaxLength(10).IsRequired();
            e.Property(x => x.DepartmentId);
            e.Property(x => x.KindCode).HasMaxLength(20);
            e.Property(x => x.DocCode).HasMaxLength(40).IsRequired();
            e.Property(x => x.DocName).HasMaxLength(200).IsRequired();
            e.Property(x => x.Title).HasMaxLength(200);
            e.Property(x => x.ApprovalCount);
            e.Property(x => x.IsActive);
            e.Property(x => x.CreatedAt);
            e.Property(x => x.CreatedBy).HasMaxLength(100);
            e.Property(x => x.UpdatedAt);
            e.Property(x => x.UpdatedBy).HasMaxLength(100);
            e.HasIndex(x => new { x.CompCd, x.DepartmentId, x.DocCode }).HasDatabaseName("IX_DocTemplateMaster_CompDeptCode");
        });

        // dbo.DocTemplateVersion
        modelBuilder.Entity<DocTemplateVersion>(e =>
        {
            e.ToTable("DocTemplateVersion", "dbo");
            e.HasKey(x => x.Id);
            e.Property(x => x.TemplateId).IsRequired();
            e.Property(x => x.VersionNo).IsRequired();
            e.Property(x => x.DescriptorJson);
            e.Property(x => x.PreviewJson);
            e.Property(x => x.Templated);
            e.Property(x => x.CreatedAt);
            e.Property(x => x.CreatedBy).HasMaxLength(100);

            // 2026.06.11 Added: 템플릿 보호 재적용 및 실제 xlsx 기준 표시 메트릭
            e.Property(x => x.PreparedAt);
            e.Property(x => x.TemplateFileHash).HasMaxLength(64).IsFixedLength();
            e.Property(x => x.ProtectionRuleCode).HasMaxLength(50);
            e.Property(x => x.VisualMetricRuleCode).HasMaxLength(50);
            e.Property(x => x.VisualSource).HasMaxLength(50);
            e.Property(x => x.VisualRangeA1).HasMaxLength(100);
            e.Property(x => x.VisualWidthPx);
            e.Property(x => x.VisualHeightPx);

            e.HasIndex(x => new { x.TemplateId, x.VersionNo }).HasDatabaseName("IX_DocTemplateVersion_TmplVer");
        });

        // dbo.DocTemplateFile
        modelBuilder.Entity<DocTemplateFile>(e =>
        {
            e.ToTable("DocTemplateFile", "dbo");
            e.HasKey(x => x.Id);
            e.Property(x => x.VersionId).IsRequired();
            e.Property(x => x.FileRole).HasMaxLength(50).IsRequired();
            e.Property(x => x.Storage).HasMaxLength(20).IsRequired();
            e.Property(x => x.FileName).HasMaxLength(255);
            e.Property(x => x.FilePath).HasMaxLength(500);
            e.Property(x => x.FileSize);
            e.Property(x => x.FileSizeBytes);
            e.Property(x => x.ContentType).HasMaxLength(200);
            e.Property(x => x.Contents);             // NVARCHAR(MAX) 가정
            e.Property(x => x.CreatedAt);
            e.Property(x => x.CreatedBy).HasMaxLength(100);
            e.HasIndex(x => new { x.VersionId, x.FileRole }).HasDatabaseName("IX_DocTemplateFile_Version_Role");
        });

        // dbo.DocTemplateApproval (옵션)
        modelBuilder.Entity<DocTemplateApproval>(e =>
        {
            e.ToTable("DocTemplateApproval", "dbo");
            e.HasKey(x => x.Id);
            e.Property(x => x.VersionId).IsRequired();
            e.Property(x => x.Slot);
            e.Property(x => x.Part).HasMaxLength(50).IsRequired();
            e.Property(x => x.A1).HasMaxLength(50);
            e.Property(x => x.Row);
            e.Property(x => x.Column);
            e.Property(x => x.CellA1).HasMaxLength(50);
            e.Property(x => x.CellRow);
            e.Property(x => x.CellColumn);
        });

        // 2026.06.16 Added: 전결 규칙 테이블 매핑 Contents 템플릿 버전별 전결권자 차수와 생략 대상 차수 및 조건 유형을 매핑
        modelBuilder.Entity<DocTemplateDelegationRule>(e =>
        {
            e.ToTable("DocTemplateDelegationRule", "dbo");
            e.HasKey(x => x.Id);

            e.Property(x => x.TemplateId).IsRequired();
            e.Property(x => x.TemplateVersionId).IsRequired();
            e.Property(x => x.RuleName).HasMaxLength(100);
            e.Property(x => x.ConditionType).IsRequired().HasMaxLength(30).IsUnicode(false);
            e.Property(x => x.DelegationStepOrder).IsRequired();
            e.Property(x => x.SkipFromStepOrder).IsRequired();
            e.Property(x => x.SkipToStepOrder).IsRequired();
            e.Property(x => x.Priority).HasDefaultValue(100);
            e.Property(x => x.IsActive).HasDefaultValue(true);
            e.Property(x => x.Note).HasMaxLength(500);
            e.Property(x => x.CreatedBy).HasMaxLength(100).IsRequired();
            e.Property(x => x.CreatedAt).HasDefaultValueSql("SYSUTCDATETIME()");
            e.Property(x => x.UpdatedBy).HasMaxLength(100);
            e.Property(x => x.UpdatedAt);
        });

        // 2026.06.16 Added: 전결 금액 조건 테이블 매핑 Contents 금액 조건 전결의 금액 필드와 통화 필드 및 통화별 기준 금액을 매핑
        modelBuilder.Entity<DocTemplateDelegationAmountRule>(e =>
        {
            e.ToTable("DocTemplateDelegationAmountRule", "dbo");
            e.HasKey(x => x.Id);

            e.Property(x => x.RuleId).IsRequired();
            e.Property(x => x.AmountFieldKey).HasMaxLength(100).IsRequired();
            e.Property(x => x.CurrencyFieldKey).HasMaxLength(100).IsRequired();

            // 2026.06.18 Added: 전결 금액 조건 셀 직접 참조 매핑 추가 Contents 입력 필드가 아닌 수식 셀과 통화 셀을 참조하기 위한 셀 주소를 매핑
            e.Property(x => x.AmountCellA1).HasMaxLength(50);
            e.Property(x => x.CurrencyCellA1).HasMaxLength(50);

            e.Property(x => x.CurrencyCode).HasMaxLength(10).IsRequired();
            e.Property(x => x.LimitAmount).HasPrecision(19, 4).IsRequired();
            e.Property(x => x.IsActive).HasDefaultValue(true);
            e.Property(x => x.CreatedBy).HasMaxLength(100).IsRequired();
            e.Property(x => x.CreatedAt).HasDefaultValueSql("SYSUTCDATETIME()");
            e.Property(x => x.UpdatedBy).HasMaxLength(100);
            e.Property(x => x.UpdatedAt);
        });

        // 2026.06.16 Added: 문서별 전결 적용 결과 테이블 매핑 Contents 문서 작성 및 승인 시점의 전결 후보와 적용 결과 스냅샷을 매핑
        modelBuilder.Entity<DocumentDelegationResult>(e =>
        {
            e.ToTable("DocumentDelegationResult", "dbo");
            e.HasKey(x => x.Id);

            e.Property(x => x.DocId).HasMaxLength(40).IsRequired().IsUnicode(false);
            e.Property(x => x.RuleId);
            e.Property(x => x.TemplateVersionId);
            e.Property(x => x.ConditionType).HasMaxLength(30).IsRequired().IsUnicode(false);
            e.Property(x => x.DelegationStepOrder).IsRequired();
            e.Property(x => x.SkipFromStepOrder).IsRequired();
            e.Property(x => x.SkipToStepOrder).IsRequired();
            e.Property(x => x.AmountFieldKey).HasMaxLength(100);
            e.Property(x => x.AmountValue).HasPrecision(19, 4);
            e.Property(x => x.CurrencyFieldKey).HasMaxLength(100);

            // 2026.06.18 Added: 문서별 전결 금액 조건 셀 직접 참조 매핑 추가 Contents 실제 문서 파일에서 비교한 금액 셀과 통화 셀 주소를 매핑
            e.Property(x => x.AmountCellA1).HasMaxLength(50);
            e.Property(x => x.CurrencyCellA1).HasMaxLength(50);

            e.Property(x => x.CurrencyCode).HasMaxLength(10);
            e.Property(x => x.LimitAmount).HasPrecision(19, 4);
            e.Property(x => x.AppliedStatus).HasMaxLength(30).IsRequired().IsUnicode(false).HasDefaultValue("Candidate");
            e.Property(x => x.AppliedBy).HasMaxLength(64);
            e.Property(x => x.AppliedAt);
            e.Property(x => x.CancelledBy).HasMaxLength(64);
            e.Property(x => x.CancelledAt);
            e.Property(x => x.ResultMessageKey).HasMaxLength(100);
            e.Property(x => x.DetailJson);
            e.Property(x => x.CreatedAt).HasDefaultValueSql("SYSUTCDATETIME()");
            e.Property(x => x.UpdatedAt);
        });

        modelBuilder.Entity<WebPushSubscription>(e =>
        {
            e.ToTable("WebPushSubscriptions");
            e.HasKey(x => x.Id);

            e.Property(x => x.UserId).HasMaxLength(64);
            e.Property(x => x.Email).HasMaxLength(256);

            e.Property(x => x.Endpoint).HasMaxLength(2048).IsRequired();
            e.Property(x => x.P256dh).HasMaxLength(256).IsRequired();
            e.Property(x => x.Auth).HasMaxLength(256).IsRequired();

            e.Property(x => x.UserAgent).HasMaxLength(512);

            e.Property(x => x.IsActive).HasDefaultValue(true);

            // DB 기본값 SYSUTCDATETIME()가 이미 있으므로, EF는 값만 보관해도 됨
            // (필요 시 ValueGeneratedOnAddOrUpdate 등을 추가할 수 있으나 현재는 최소 변경)
        });
    }
}