// File: Data/Config/TemplateKindMasterLocConfig.cs
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using WebApplication1.Models;

namespace WebApplication1.Data.Config
{
    public sealed class TemplateKindMasterLocConfig : IEntityTypeConfiguration<TemplateKindMasterLoc>
    {
        public void Configure(EntityTypeBuilder<TemplateKindMasterLoc> e)
        {
            e.ToTable("TemplateKindMasterLoc", "dbo");

            // ▼ PK: (Id, LangCode) — Id는 TemplateKindMaster.Id FK
            e.HasKey(x => new { x.Id, x.LangCode });

            e.Property(x => x.Id)
             .IsRequired();

            e.Property(x => x.LangCode)
             .HasMaxLength(10)
             .IsRequired();

            // ▼ 스키마 정합: Name 64자로 맞춤
            e.Property(x => x.Name)
             .HasMaxLength(64)
             .IsRequired();

            // ▼ rowversion(동시성) — 그림자 속성
            e.Property<byte[]>("RowVersion").IsRowVersion();

            // 조회 편의 인덱스
            e.HasIndex(x => x.LangCode);

            // FK: TemplateKindMaster(Id) ← TemplateKindMasterLoc(Id)
            e.HasOne<TemplateKindMaster>()
             .WithMany() // 네비게이션을 쓰려면 .WithMany(m => m.Locs)
             .HasForeignKey(x => x.Id)
             .OnDelete(DeleteBehavior.Cascade);

            // (선택) 유효성 보조 체크
            e.ToTable(tb =>
            {
                tb.HasCheckConstraint("CK_TemplateKindMasterLoc_Name", "LEN(LTRIM(RTRIM([Name]))) > 0");
            });
        }
    }
}
