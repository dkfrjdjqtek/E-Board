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
