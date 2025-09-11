using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    public partial class Update_TemplateKind_Schema_20250910 : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"
SET XACT_ABORT ON;
BEGIN TRAN;

------------------------------------------------------------
-- TemplateKindMasterLoc : 스키마 정합화 (필요시 교체 생성)
------------------------------------------------------------
IF OBJECT_ID(N'[dbo].[TemplateKindMasterLoc]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[TemplateKindMasterLoc](
        [Id]        INT           NOT NULL,           -- FK → TemplateKindMasters.Id
        [LangCode]  NVARCHAR(10)  NOT NULL,           -- 길이 10
        [Name]      NVARCHAR(100) NOT NULL,           -- 길이 100
        [RowVersion] ROWVERSION    NOT NULL,
        CONSTRAINT [PK_TemplateKindMasterLoc] PRIMARY KEY ([Id],[LangCode]),
        CONSTRAINT [FK_TemplateKindMasterLoc_TemplateKindMasters_Id]
            FOREIGN KEY ([Id]) REFERENCES [dbo].[TemplateKindMasters]([Id]) ON DELETE CASCADE
    );
    CREATE INDEX [IX_TemplateKindMasterLoc_LangCode] ON [dbo].[TemplateKindMasterLoc]([LangCode]);

    IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_TemplateKindMasterLoc_Name')
        ALTER TABLE [dbo].[TemplateKindMasterLoc] WITH NOCHECK
        ADD CONSTRAINT [CK_TemplateKindMasterLoc_Name] CHECK (LEN(LTRIM(RTRIM([Name]))) > 0);
END
ELSE
BEGIN
    -- 원하는 규격과 다를 수 있으므로 새 테이블로 교체 후 데이터 이관
    CREATE TABLE [dbo].[__TemplateKindMasterLoc_new](
        [Id]        INT           NOT NULL,
        [LangCode]  NVARCHAR(10)  NOT NULL,
        [Name]      NVARCHAR(100) NOT NULL,
        [RowVersion] ROWVERSION    NOT NULL,
        CONSTRAINT [PK_TemplateKindMasterLoc] PRIMARY KEY ([Id],[LangCode]),
        CONSTRAINT [FK_TemplateKindMasterLoc_TemplateKindMasters_Id]
            FOREIGN KEY ([Id]) REFERENCES [dbo].[TemplateKindMasters]([Id]) ON DELETE CASCADE
    );

    INSERT INTO [dbo].[__TemplateKindMasterLoc_new]([Id],[LangCode],[Name])
    SELECT [Id],
           LEFT(COALESCE([LangCode], N''), 10),
           LEFT(COALESCE([Name], N''), 100)
    FROM   [dbo].[TemplateKindMasterLoc] WITH (HOLDLOCK);

    DROP TABLE [dbo].[TemplateKindMasterLoc];
    EXEC sp_rename N'[dbo].[__TemplateKindMasterLoc_new]', N'TemplateKindMasterLoc';

    CREATE INDEX [IX_TemplateKindMasterLoc_LangCode] ON [dbo].[TemplateKindMasterLoc]([LangCode]);

    IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_TemplateKindMasterLoc_Name')
        ALTER TABLE [dbo].[TemplateKindMasterLoc] WITH NOCHECK
        ADD CONSTRAINT [CK_TemplateKindMasterLoc_Name] CHECK (LEN(LTRIM(RTRIM([Name]))) > 0);
END

------------------------------------------------------------
-- TemplateKindMasters : 누락 컬럼/인덱스/체크 제약 보장
--  (Id의 IDENTITY 전환은 별도 마이그레이션에서 처리됨)
------------------------------------------------------------
IF OBJECT_ID(N'[dbo].[TemplateKindMasters]', N'U') IS NOT NULL
BEGIN
    -- SortOrder 없으면 추가
    IF COL_LENGTH('dbo.TemplateKindMasters','SortOrder') IS NULL
        ALTER TABLE [dbo].[TemplateKindMasters]
        ADD [SortOrder] INT NOT NULL CONSTRAINT [DF_TemplateKindMasters_SortOrder_20250910] DEFAULT(0) WITH VALUES;

    -- RowVersion 없으면 추가
    IF COL_LENGTH('dbo.TemplateKindMasters','RowVersion') IS NULL
        ALTER TABLE [dbo].[TemplateKindMasters]
        ADD [RowVersion] ROWVERSION NOT NULL;

    -- (CompCd, Code) 유니크 인덱스
    IF NOT EXISTS (
        SELECT 1 FROM sys.indexes
        WHERE name = N'IX_TemplateKindMasters_CompCd_Code'
          AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
    )
        CREATE UNIQUE INDEX [IX_TemplateKindMasters_CompCd_Code]
        ON [dbo].[TemplateKindMasters]([CompCd],[Code]);

    -- (CompCd, DepartmentId) 조회용 인덱스
    IF NOT EXISTS (
        SELECT 1 FROM sys.indexes
        WHERE name = N'IX_TemplateKindMasters_CompCd_Department'
          AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
    )
        CREATE INDEX [IX_TemplateKindMasters_CompCd_Department]
        ON [dbo].[TemplateKindMasters]([CompCd],[DepartmentId]);

    -- Code 공백 금지 체크 제약
    IF NOT EXISTS (
        SELECT 1 FROM sys.check_constraints
        WHERE name = N'CK_TemplateKindMasters_Code'
          AND parent_object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
    )
        ALTER TABLE [dbo].[TemplateKindMasters] WITH NOCHECK
        ADD CONSTRAINT [CK_TemplateKindMasters_Code] CHECK (LEN(LTRIM(RTRIM([Code]))) > 0);
END

COMMIT TRAN;
");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"
SET XACT_ABORT ON;
BEGIN TRAN;

-- 본 마이그레이션에서 생성한 인덱스/체크 제약만 제거 (스키마 원복은 최소화)
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

    IF EXISTS (
        SELECT 1 FROM sys.check_constraints
        WHERE name = N'CK_TemplateKindMasters_Code'
          AND parent_object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
    )
        ALTER TABLE [dbo].[TemplateKindMasters] DROP CONSTRAINT [CK_TemplateKindMasters_Code];

    -- 추가했던 기본값 제약 이름 안전 제거 (존재 시)
    IF COL_LENGTH('dbo.TemplateKindMasters','SortOrder') IS NOT NULL
    BEGIN
        DECLARE @df NVARCHAR(128);
        SELECT @df = d.name
        FROM sys.default_constraints d
        JOIN sys.columns c ON c.default_object_id = d.object_id
        WHERE d.parent_object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')
          AND c.name = N'SortOrder';
        IF @df IS NOT NULL EXEC('ALTER TABLE [dbo].[TemplateKindMasters] DROP CONSTRAINT [' + @df + ']');
        -- 컬럼 자체는 남깁니다(데이터 보존)
    END
END

IF OBJECT_ID(N'[dbo].[TemplateKindMasterLoc]', N'U') IS NOT NULL
BEGIN
    IF EXISTS (
        SELECT 1 FROM sys.indexes
        WHERE name = N'IX_TemplateKindMasterLoc_LangCode'
          AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasterLoc]')
    )
        DROP INDEX [IX_TemplateKindMasterLoc_LangCode] ON [dbo].[TemplateKindMasterLoc];

    IF EXISTS (
        SELECT 1 FROM sys.check_constraints
        WHERE name = N'CK_TemplateKindMasterLoc_Name'
          AND parent_object_id = OBJECT_ID(N'[dbo].[TemplateKindMasterLoc]')
    )
        ALTER TABLE [dbo].[TemplateKindMasterLoc] DROP CONSTRAINT [CK_TemplateKindMasterLoc_Name];
END

COMMIT TRAN;
");
        }
    }
}
