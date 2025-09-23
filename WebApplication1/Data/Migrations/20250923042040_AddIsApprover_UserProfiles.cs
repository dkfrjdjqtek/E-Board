// 2025.09.23 Changed: 기존 객체가 있을 때 마이그레이션 충돌 방지 위해 조건부 DDL로 교체 테이블/인덱스 IF NOT EXISTS, 컬럼/제약 동적 처리 (FK는 생성하지 않음)
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    public partial class AddIsApprover_UserProfiles : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            // UserProfiles.IsApprover: 이미 있으면 스킵
            migrationBuilder.Sql(@"
IF COL_LENGTH('dbo.UserProfiles','IsApprover') IS NULL
BEGIN
    ALTER TABLE dbo.UserProfiles 
        ADD IsApprover BIT NOT NULL CONSTRAINT DF_UserProfiles_IsApprover DEFAULT(0);
END
");

            // TemplateKindMasters: 이미 있으면 스킵
            migrationBuilder.Sql(@"
IF OBJECT_ID('dbo.TemplateKindMasters','U') IS NULL
BEGIN
    CREATE TABLE dbo.TemplateKindMasters
    (
        Id          INT IDENTITY(1,1) NOT NULL,
        CompCd      NVARCHAR(10)  NOT NULL,
        DepartmentId INT          NOT NULL CONSTRAINT DF_TemplateKindMasters_DepartmentId DEFAULT(0),
        Code        NVARCHAR(32)  NOT NULL,
        Name        NVARCHAR(64)  NOT NULL,
        IsActive    BIT           NOT NULL CONSTRAINT DF_TemplateKindMasters_IsActive DEFAULT(1),
        SortOrder   INT           NOT NULL CONSTRAINT DF_TemplateKindMasters_SortOrder DEFAULT(0),
        RowVersion  ROWVERSION    NULL,
        CONSTRAINT PK_TemplateKindMasters PRIMARY KEY (Id)
    );
END;

-- 인덱스들: 없을 때만 생성
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_TemplateKindMasters_CompCd_Department' AND object_id = OBJECT_ID('dbo.TemplateKindMasters'))
    CREATE INDEX IX_TemplateKindMasters_CompCd_Department ON dbo.TemplateKindMasters(CompCd, DepartmentId);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UX_TemplateKindMasters_CompCd_Code' AND object_id = OBJECT_ID('dbo.TemplateKindMasters'))
    CREATE UNIQUE INDEX UX_TemplateKindMasters_CompCd_Code ON dbo.TemplateKindMasters(CompCd, Code);
");

            // TemplateKindMasterLoc: 이미 있으면 스킵 (주의: FK 생성하지 않음 - 프로젝트 규칙)
            migrationBuilder.Sql(@"
IF OBJECT_ID('dbo.TemplateKindMasterLoc','U') IS NULL
BEGIN
    CREATE TABLE dbo.TemplateKindMasterLoc
    (
        Id          INT           NOT NULL,
        LangCode    NVARCHAR(10)  NOT NULL,
        CompCd      NVARCHAR(10)  NOT NULL,
        DepartmentId INT          NOT NULL CONSTRAINT DF_TemplateKindMasterLoc_DepartmentId DEFAULT(0),
        Name        NVARCHAR(64)  NOT NULL,
        RowVersion  ROWVERSION    NULL,
        CONSTRAINT PK_TemplateKindMasterLoc PRIMARY KEY (Id, LangCode)
    );
END;

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_TemplateKindMasterLoc_CompCd_DepartmentId_LangCode' AND object_id = OBJECT_ID('dbo.TemplateKindMasterLoc'))
    CREATE INDEX IX_TemplateKindMasterLoc_CompCd_DepartmentId_LangCode ON dbo.TemplateKindMasterLoc(CompCd, DepartmentId, LangCode);
");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            // TemplateKindMasterLoc: 존재할 때만 드롭
            migrationBuilder.Sql(@"
IF OBJECT_ID('dbo.TemplateKindMasterLoc','U') IS NOT NULL
BEGIN
    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_TemplateKindMasterLoc_CompCd_DepartmentId_LangCode' AND object_id = OBJECT_ID('dbo.TemplateKindMasterLoc'))
        DROP INDEX IX_TemplateKindMasterLoc_CompCd_DepartmentId_LangCode ON dbo.TemplateKindMasterLoc;
    DROP TABLE dbo.TemplateKindMasterLoc;
END
");

            // TemplateKindMasters: 존재할 때만 드롭
            migrationBuilder.Sql(@"
IF OBJECT_ID('dbo.TemplateKindMasters','U') IS NOT NULL
BEGIN
    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UX_TemplateKindMasters_CompCd_Code' AND object_id = OBJECT_ID('dbo.TemplateKindMasters'))
        DROP INDEX UX_TemplateKindMasters_CompCd_Code ON dbo.TemplateKindMasters;
    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_TemplateKindMasters_CompCd_Department' AND object_id = OBJECT_ID('dbo.TemplateKindMasters'))
        DROP INDEX IX_TemplateKindMasters_CompCd_Department ON dbo.TemplateKindMasters;
    DROP TABLE dbo.TemplateKindMasters;
END
");

            // UserProfiles.IsApprover: 존재할 때만 드롭(기본제약 동적 제거)
            migrationBuilder.Sql(@"
IF COL_LENGTH('dbo.UserProfiles','IsApprover') IS NOT NULL
BEGIN
    DECLARE @df sysname;
    SELECT @df = dc.name
    FROM sys.default_constraints dc
    JOIN sys.columns c ON c.default_object_id = dc.object_id
    JOIN sys.tables t ON t.object_id = c.object_id
    WHERE t.name = 'UserProfiles'
      AND SCHEMA_NAME(t.schema_id) = 'dbo'
      AND c.name = 'IsApprover';

    IF @df IS NOT NULL
        EXEC('ALTER TABLE dbo.UserProfiles DROP CONSTRAINT ' + QUOTENAME(@df));

    ALTER TABLE dbo.UserProfiles DROP COLUMN IsApprover;
END
");
        }
    }
}
