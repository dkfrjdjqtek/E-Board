using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using WebApplication1.Models;

namespace WebApplication1.Data;

public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
{
    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options) { }

    public DbSet<WebAuthnCredential> WebAuthnCredentials => Set<WebAuthnCredential>();
    public DbSet<UserProfile> UserProfiles => Set<UserProfile>();
    public DbSet<DepartmentMaster> DepartmentMasters => Set<DepartmentMaster>();
    public DbSet<PositionMaster> PositionMasters => Set<PositionMaster>();
    public DbSet<DepartmentMasterLoc> DepartmentMasterLocs => Set<DepartmentMasterLoc>();
    public DbSet<PositionMasterLoc> PositionMasterLocs => Set<PositionMasterLoc>();

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
            e.HasIndex(x => x.UserId).IsUnique();
            e.Property(x => x.CompCd).HasMaxLength(10).IsRequired().HasConversion(UpperTrim);
            e.Property(x => x.DisplayName).HasMaxLength(64);

            e.HasOne(p => p.User)
             .WithOne(u => u.Profile)
             .HasForeignKey<UserProfile>(p => p.UserId)
             .OnDelete(DeleteBehavior.Cascade);

            e.HasOne(p => p.Department).WithMany()
             .HasForeignKey(p => p.DepartmentId)
             .OnDelete(DeleteBehavior.SetNull);

            e.HasOne(p => p.Position).WithMany()
             .HasForeignKey(p => p.PositionId)
             .OnDelete(DeleteBehavior.SetNull);

            e.Property<byte[]>("RowVersion").IsRowVersion();
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
    }
}
