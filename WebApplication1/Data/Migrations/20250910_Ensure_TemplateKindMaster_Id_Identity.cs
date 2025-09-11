using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    public partial class Ensure_TemplateKindMaster_Id_Identity : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"
SET XACT_ABORT ON;
BEGIN TRAN;

-- TemplateKindMasters 가 없으면 초기 구축 단계이므로 스킵
IF OBJECT_ID(N'[dbo].[TemplateKindMasters]', N'U') IS NULL
BEGIN
    COMMIT TRAN;
    RETURN;
END

DECLARE @is_identity INT =
    (SELECT COLUMNPROPERTY(OBJECT_ID(N'[dbo].[TemplateKindMasters]'), 'Id', 'IsIdentity'));

-- 공통: FK/인덱스/체크제약 존재 여부 캐시
DECLARE @has_fk_loc BIT =
    CASE WHEN EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_TemplateKindMasterLoc_TemplateKindMasters_Id') THEN 1 ELSE 0 END;

DECLARE @has_ix_code BIT =
    CASE WHEN EXISTS (SELECT 1 FROM sys.indexes 
                      WHERE name = N'IX_TemplateKindMasters_CompCd_Code'
                      AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')) THEN 1 ELSE 0 END;

DECLARE @has_ix_dept BIT =
    CASE WHEN EXISTS (SELECT 1 FROM sys.indexes 
                      WHERE name = N'IX_TemplateKindMasters_CompCd_Department'
                      AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')) THEN 1 ELSE 0 END;

DECLARE @has_ck_code BIT =
    CASE WHEN EXISTS (SELECT 1 FROM sys.check_constraints 
                      WHERE name = N'CK_TemplateKindMasters_Code'
                      AND parent_object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]')) THEN 1 ELSE 0 END;

IF (@is_identity = 0)
BEGIN
    -- 1) FK/인덱스/체크 제약 제거
    IF (@has_fk_loc = 1)
        ALTER TABLE [dbo].[TemplateKindMasterLoc] DROP CONSTRAINT [FK_TemplateKindMasterLoc_TemplateKindMasters_Id];

    IF (@has_ix_code = 1)
        DROP INDEX [IX_TemplateKindMasters_CompCd_Code] ON [dbo].[TemplateKindMasters];

    IF (@has_ix_dept = 1)
        DROP INDEX [IX_TemplateKindMasters_CompCd_Department] ON [dbo].[TemplateKindMasters];

    IF (@has_ck_code = 1)
        ALTER TABLE [dbo].[TemplateKindMasters] DROP CONSTRAINT [CK_TemplateKindMasters_Code];

    -- 2) 새 테이블(IDENTITY, 스키마 정합) 생성
    CREATE TABLE [dbo].[__TemplateKindMasters_new](
        [Id]           INT IDENTITY(1,1) NOT NULL,
        [CompCd]       NVARCHAR(10)  NOT NULL,
        [DepartmentId] INT           NOT NULL CONSTRAINT [DF___Tkm_DepartmentId] DEFAULT (0),
        [Code]         NVARCHAR(32)  NOT NULL,
        [Name]         NVARCHAR(64)  NOT NULL,
        [IsActive]     BIT           NOT NULL CONSTRAINT [DF___Tkm_IsActive] DEFAULT (1),
        [SortOrder]    INT           NOT NULL CONSTRAINT [DF___Tkm_SortOrder] DEFAULT (0),
        [RowVersion]   rowversion    NOT NULL,
        CONSTRAINT [PK___TemplateKindMasters_new] PRIMARY KEY CLUSTERED ([Id] ASC)
    );

    -- 3) 데이터 이관 (Id 유지 위해 IDENTITY_INSERT)
    SET IDENTITY_INSERT [dbo].[__TemplateKindMasters_new] ON;

    INSERT INTO [dbo].[__TemplateKindMasters_new]
        ([Id],[CompCd],[DepartmentId],[Code],[Name],[IsActive],[SortOrder])
    SELECT
        [Id],[CompCd],[DepartmentId],[Code],[Name],[IsActive],
        ISNULL([SortOrder], 0)
    FROM [dbo].[TemplateKindMasters] WITH (HOLDLOCK TABLOCKX);

    SET IDENTITY_INSERT [dbo].[__TemplateKindMasters_new] OFF;

    -- 4) 원본 교체
    DROP TABLE [dbo].[TemplateKindMasters];
    EXEC sp_rename N'[dbo].[__TemplateKindMasters_new]', N'TemplateKindMasters';

    -- 5) 인덱스/체크 제약 복원
    CREATE UNIQUE INDEX [IX_TemplateKindMasters_CompCd_Code]
        ON [dbo].[TemplateKindMasters]([CompCd],[Code]);

    CREATE INDEX [IX_TemplateKindMasters_CompCd_Department]
        ON [dbo].[TemplateKindMasters]([CompCd],[DepartmentId]);

    ALTER TABLE [dbo].[TemplateKindMasters]  WITH CHECK
        ADD CONSTRAINT [CK_TemplateKindMasters_Code] CHECK (LEN(LTRIM(RTRIM([Code]))) > 0);

    -- 6) FK 복원
    IF OBJECT_ID(N'[dbo].[TemplateKindMasterLoc]', N'U') IS NOT NULL
    BEGIN
        ALTER TABLE [dbo].[TemplateKindMasterLoc]  WITH CHECK
        ADD CONSTRAINT [FK_TemplateKindMasterLoc_TemplateKindMasters_Id]
            FOREIGN KEY([Id]) REFERENCES [dbo].[TemplateKindMasters]([Id]) ON DELETE CASCADE;
    END
END
ELSE IF (@is_identity = 1)
BEGIN
    -- 이미 IDENTITY: 스키마 요소 보강 (인덱스/체크/FK/누락 컬럼)
    -- 누락 컬럼 보정: SortOrder / RowVersion
    IF COL_LENGTH('dbo.TemplateKindMasters','SortOrder') IS NULL
        ALTER TABLE [dbo].[TemplateKindMasters] ADD [SortOrder] INT NOT NULL CONSTRAINT [DF___Tkm_SortOrder_Online] DEFAULT(0) WITH VALUES;

    IF COL_LENGTH('dbo.TemplateKindMasters','RowVersion') IS NULL
        ALTER TABLE [dbo].[TemplateKindMasters] ADD [RowVersion] rowversion NOT NULL;

    IF NOT EXISTS (SELECT 1 FROM sys.indexes 
                   WHERE name = N'IX_TemplateKindMasters_CompCd_Code'
                     AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]'))
        CREATE UNIQUE INDEX [IX_TemplateKindMasters_CompCd_Code]
            ON [dbo].[TemplateKindMasters]([CompCd],[Code]);

    IF NOT EXISTS (SELECT 1 FROM sys.indexes 
                   WHERE name = N'IX_TemplateKindMasters_CompCd_Department'
                     AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]'))
        CREATE INDEX [IX_TemplateKindMasters_CompCd_Department]
            ON [dbo].[TemplateKindMasters]([CompCd],[DepartmentId]);

    IF NOT EXISTS (SELECT 1 FROM sys.check_constraints 
                   WHERE name = N'CK_TemplateKindMasters_Code'
                     AND parent_object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]'))
        ALTER TABLE [dbo].[TemplateKindMasters]  WITH CHECK
            ADD CONSTRAINT [CK_TemplateKindMasters_Code] CHECK (LEN(LTRIM(RTRIM([Code]))) > 0);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys 
                   WHERE name = N'FK_TemplateKindMasterLoc_TemplateKindMasters_Id')
       AND OBJECT_ID(N'[dbo].[TemplateKindMasterLoc]', N'U') IS NOT NULL
    BEGIN
        ALTER TABLE [dbo].[TemplateKindMasterLoc]  WITH CHECK
        ADD CONSTRAINT [FK_TemplateKindMasterLoc_TemplateKindMasters_Id]
            FOREIGN KEY([Id]) REFERENCES [dbo].[TemplateKindMasters]([Id]) ON DELETE CASCADE;
    END
END
ELSE
BEGIN
    -- @is_identity IS NULL (열 조회 실패) → 아무 것도 하지 않음
END

COMMIT TRAN;
");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            // 위험한 IDENTITY 해제/원복은 하지 않습니다. 보강했던 보조 객체만 제거.
            migrationBuilder.Sql(@"
IF EXISTS (SELECT 1 FROM sys.check_constraints 
           WHERE name = N'CK_TemplateKindMasters_Code' 
             AND parent_object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]'))
BEGIN
    ALTER TABLE [dbo].[TemplateKindMasters] DROP CONSTRAINT [CK_TemplateKindMasters_Code];
END

IF EXISTS (SELECT 1 FROM sys.indexes 
           WHERE name = N'IX_TemplateKindMasters_CompCd_Department' 
             AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]'))
BEGIN
    DROP INDEX [IX_TemplateKindMasters_CompCd_Department] ON [dbo].[TemplateKindMasters];
END

IF EXISTS (SELECT 1 FROM sys.indexes 
           WHERE name = N'IX_TemplateKindMasters_CompCd_Code' 
             AND object_id = OBJECT_ID(N'[dbo].[TemplateKindMasters]'))
BEGIN
    DROP INDEX [IX_TemplateKindMasters_CompCd_Code] ON [dbo].[TemplateKindMasters];
END
");
        }
    }
}
