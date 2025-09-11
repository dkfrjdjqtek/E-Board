// File: Data/Migrations/20250910_Update_TemplateKind_Loc_PK_and_Index.cs
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    public partial class Update_TemplateKind_Loc_PK_and_Index : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"
SET XACT_ABORT ON;
BEGIN TRAN;

-- =========================================================
-- TemplateKindMasterLoc 정합화 (PK, FK, 길이, RowVersion, 인덱스)
-- =========================================================
IF OBJECT_ID(N'[dbo].[TemplateKindMasterLoc]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[TemplateKindMasterLoc](
        [Id]       INT           NOT NULL,             -- FK: TemplateKindMasters.Id
        [LangCode] NVARCHAR(10)  NOT NULL,             -- ↔ Config: .HasMaxLength(10)
        [Name]     NVARCHAR(100) NOT NULL,             -- ↔ Config: .HasMaxLength(100)
        [RowVersion] rowversion  NOT NULL,
        CONSTRAINT [PK_TemplateKindMasterLoc] PRIMARY KEY ([Id],[LangCode]),
        CONSTRAINT [FK_TemplateKindMasterLoc_TemplateKindMasters_Id]
            FOREIGN KEY([Id]) REFERENCES [dbo].[TemplateKindMasters]([Id]) ON DELETE CASCADE
    );
    CREATE INDEX [IX_TemplateKindMasterLoc_LangCode] ON [dbo].[TemplateKindMasterLoc]([LangCode]);
END
ELSE
BEGIN
    -- 새 규격 테이블 생성
    CREATE TABLE [dbo].[__TemplateKindMasterLoc_new](
        [Id]       INT           NOT NULL,
        [LangCode] NVARCHAR(10)  NOT NULL,
        [Name]     NVARCHAR(100) NOT NULL,
        [RowVersion] rowversion  NOT NULL,
        CONSTRAINT [PK_TemplateKindMasterLoc] PRIMARY KEY ([Id],[LangCode]),
        CONSTRAINT [FK_TemplateKindMasterLoc_TemplateKindMasters_Id]
            FOREIGN KEY([Id]) REFERENCES [dbo].[TemplateKindMasters]([Id]) ON DELETE CASCADE
    );

    -- 데이터 이관 (길이 초과 방지)
    INSERT INTO [dbo].[__TemplateKindMasterLoc_new]([Id],[LangCode],[Name])
    SELECT [Id], LEFT([LangCode],10), LEFT([Name],100)
    FROM   [dbo].[TemplateKindMasterLoc] WITH (HOLDLOCK);

    -- 교체
    DROP TABLE [dbo].[TemplateKindMasterLoc];
    EXEC sp_rename N'[dbo].[__TemplateKindMasterLoc_new]', N'TemplateKindMasterLoc';

    -- 인덱스 복원
    CREATE INDEX [IX_TemplateKindMasterLoc_LangCode] ON [dbo].[TemplateKindMasterLoc]([LangCode]);
END

-- =========================================================
-- TemplateKindMasters 보강 (인덱스/누락컬럼만)
--   ※ 테이블 스키마 자체 변경은 다른 마이그레이션에서 수행됨
-- =========================================================
IF OBJECT_ID(N'[dbo].[TemplateKindMasters]', N'U') IS NOT NULL
BEGIN
    -- 누락 컬럼 보정
    IF COL_LENGTH('dbo.TemplateKindMasters','SortOrder') IS NULL
        ALTER TABLE [dbo].[TemplateKindMasters] ADD [SortOrder] INT NOT NULL CONSTRAINT [DF_TKM_SortOrder_20250910] DEFAULT(0) WITH VALUES;

    IF COL_LENGTH('dbo.TemplateKindMasters','RowVersion') IS NULL
        ALTER TABLE [dbo].[TemplateKindMasters] ADD [RowVersion] rowversion NOT NULL;

    -- (CompCd, Code) 유니크
    IF NOT EXISTS (
        SELECT 1 FROM sys.indexes
        WHERE name = N'IX_TemplateKindMasters_CompCd_Code'
          AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
    )
    BEGIN
        CREATE UNIQUE INDEX [IX_TemplateKindMasters_CompCd_Code]
        ON [dbo].[TemplateKindMasters]([CompCd],[Code]);
    END

    -- (CompCd, DepartmentId) 검색용
    IF NOT EXISTS (
        SELECT 1 FROM sys.indexes
        WHERE name = N'IX_TemplateKindMasters_CompCd_Department'
          AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
    )
    BEGIN
        CREATE INDEX [IX_TemplateKindMasters_CompCd_Department]
        ON [dbo].[TemplateKindMasters]([CompCd],[DepartmentId]);
    END
END

COMMIT TRAN;
");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"
SET XACT_ABORT ON;
BEGIN TRAN;

-- 인덱스 롤백(본 마이그레이션에서 만든 것만)
IF OBJECT_ID(N'[dbo].[TemplateKindMasters]', N'U') IS NOT NULL
BEGIN
    IF EXISTS (
        SELECT 1 FROM sys.indexes
        WHERE name = N'IX_TemplateKindMasters_CompCd_Department'
          AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
    )
        DROP INDEX [IX_TemplateKindMasters_CompCd_Department] ON [dbo].[TemplateKindMasters];

    IF EXISTS (
        SELECT 1 FROM sys.indexes
        WHERE name = N'IX_TemplateKindMasters_CompCd_Code'
          AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
    )
        DROP INDEX [IX_TemplateKindMasters_CompCd_Code] ON [dbo].[TemplateKindMasters];
END

-- TemplateKindMasterLoc: 단순 제거(Up에서 교체 생성했으므로)
IF OBJECT_ID(N'[dbo].[TemplateKindMasterLoc]', N'U') IS NOT NULL
BEGIN
    DROP TABLE [dbo].[TemplateKindMasterLoc];
END

COMMIT TRAN;
");
        }
    }
}
