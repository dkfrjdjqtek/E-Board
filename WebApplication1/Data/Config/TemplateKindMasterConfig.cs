// File: Data/Config/TemplateKindMasterConfig.cs
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using WebApplication1.Models;

namespace WebApplication1.Data.Config
{
    public sealed class TemplateKindMasterConfig : IEntityTypeConfiguration<TemplateKindMaster>
    {
        public void Configure(EntityTypeBuilder<TemplateKindMaster> e)
        {
            e.ToTable("TemplateKindMasters", "dbo");

            e.HasKey(x => x.Id);
            e.Property(x => x.Id)
             .ValueGeneratedOnAdd(); // INT IDENTITY(1,1)

            e.Property(x => x.CompCd)
             .HasMaxLength(10)
             .IsRequired();

            e.Property(x => x.DepartmentId)
             .HasDefaultValue(0);

            e.Property(x => x.Code)
             .HasMaxLength(32)
             .IsRequired();

            // DB 스키마에 맞춰 64자로 축소
            e.Property(x => x.Name)
             .HasMaxLength(64)
             .IsRequired();

            e.Property(x => x.IsActive)
             .HasDefaultValue(true);

            // ▼ 스키마 정합: SortOrder 컬럼(모델에는 없지만 DB에는 필요) — 기본값 0
            e.Property<int>("SortOrder").HasDefaultValue(0);

            // ▼ rowversion(동시성 토큰) — 그림자 속성으로 매핑
            e.Property<byte[]>("RowVersion").IsRowVersion();

            // comp별 코드 유니크 (T0001, T0002 …)
            e.HasIndex(x => new { x.CompCd, x.Code })
             .IsUnique();

            // 조회 최적화용(사업장/부서)
            e.HasIndex(x => new { x.CompCd, x.DepartmentId });

            e.ToTable(tb =>
            {
                tb.HasCheckConstraint("CK_TemplateKindMasters_Code", "LEN(LTRIM(RTRIM([Code]))) > 0");
            });

            // (선택) 네비게이션을 쓰는 경우에만 활성화
            // e.HasMany(x => x.Locs)
            //  .WithOne(l => l.Master)
            //  .HasForeignKey(l => l.Id)
            //  .OnDelete(DeleteBehavior.Cascade);
        }
    }
}
