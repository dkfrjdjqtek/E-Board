using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    /// <inheritdoc />
    public partial class Add_PositionMasterLoc : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "Discriminator",
                table: "AspNetUsers",
                type: "nvarchar(21)",
                maxLength: 21,
                nullable: false,
                defaultValue: "");

            migrationBuilder.CreateTable(
                name: "DepartmentMasters",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CompCd = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    Code = table.Column<string>(type: "nvarchar(32)", maxLength: 32, nullable: false),
                    Name = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: false),
                    IsActive = table.Column<bool>(type: "bit", nullable: false),
                    SortOrder = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DepartmentMasters", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "PositionMasters",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CompCd = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    Code = table.Column<string>(type: "nvarchar(32)", maxLength: 32, nullable: false),
                    Name = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: false),
                    IsActive = table.Column<bool>(type: "bit", nullable: false),
                    RankLevel = table.Column<short>(type: "smallint", nullable: false, defaultValue: (short)0),
                    IsApprover = table.Column<bool>(type: "bit", nullable: false, defaultValue: false),
                    SortOrder = table.Column<int>(type: "int", nullable: false, defaultValue: 0)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PositionMasters", x => x.Id);
                });

            migrationBuilder.Sql(@"
-- 1) 테이블 없으면 생성
IF OBJECT_ID(N'dbo.WebAuthnCredentials','U') IS NULL
BEGIN
    CREATE TABLE [dbo].[WebAuthnCredentials](
        [Id] UNIQUEIDENTIFIER NOT NULL,
        [UserId] NVARCHAR(450) NOT NULL,
        [CredentialId] VARBINARY(MAX) NOT NULL,
        [CredentialIdHash] VARBINARY(900) NOT NULL CONSTRAINT DF_WebAuthn_CredHash DEFAULT 0x,
        [PublicKey] VARBINARY(MAX) NOT NULL,
        [CredType] NVARCHAR(20) NOT NULL CONSTRAINT DF_WebAuthn_CredType DEFAULT N'public-key',
        [UserHandle] VARBINARY(MAX) NULL,
        [SignCount] INT NOT NULL CONSTRAINT DF_WebAuthn_SignCount DEFAULT(0),
        [IsBackupEligible] BIT NOT NULL CONSTRAINT DF_WebAuthn_BackupEligible DEFAULT(0),
        [IsBackedUp] BIT NOT NULL CONSTRAINT DF_WebAuthn_BackedUp DEFAULT(0),
        [Transports] NVARCHAR(200) NULL,
        [Nickname] NVARCHAR(100) NULL,
        [CreatedAtUtc] DATETIME2 NOT NULL CONSTRAINT DF_WebAuthn_CreatedAt DEFAULT SYSUTCDATETIME(),
        [LastUsedAtUtc] DATETIME2 NULL,
        [IsDiscoverable] BIT NOT NULL CONSTRAINT DF_WebAuthn_IsDiscoverable DEFAULT(0),
        [AaGuid] UNIQUEIDENTIFIER NULL,
        CONSTRAINT [PK_WebAuthnCredentials] PRIMARY KEY ([Id])
    );
END;

-- 2) 누락된 컬럼만 추가 (형식 변경/ALTER COLUMN 금지)
IF COL_LENGTH('dbo.WebAuthnCredentials','CredentialIdHash') IS NULL
    ALTER TABLE dbo.WebAuthnCredentials ADD CredentialIdHash VARBINARY(900) NOT NULL CONSTRAINT DF_WebAuthn_CredHash DEFAULT 0x;

IF COL_LENGTH('dbo.WebAuthnCredentials','IsDiscoverable') IS NULL
    ALTER TABLE dbo.WebAuthnCredentials ADD IsDiscoverable BIT NOT NULL CONSTRAINT DF_WebAuthn_IsDiscoverable DEFAULT(0);

IF COL_LENGTH('dbo.WebAuthnCredentials','AaGuid') IS NULL
    ALTER TABLE dbo.WebAuthnCredentials ADD AaGuid UNIQUEIDENTIFIER NULL;

-- 3) 인덱스/제약조건 없을 때만 생성
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_WebAuthnCredentials_CredentialIdHash' AND object_id=OBJECT_ID('dbo.WebAuthnCredentials'))
    CREATE UNIQUE INDEX IX_WebAuthnCredentials_CredentialIdHash ON dbo.WebAuthnCredentials(CredentialIdHash);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_WebAuthnCredentials_UserId' AND object_id=OBJECT_ID('dbo.WebAuthnCredentials'))
    CREATE INDEX IX_WebAuthnCredentials_UserId ON dbo.WebAuthnCredentials(UserId);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_WebAuthnCredentials_UserId_Nickname' AND object_id=OBJECT_ID('dbo.WebAuthnCredentials'))
    CREATE UNIQUE INDEX IX_WebAuthnCredentials_UserId_Nickname ON dbo.WebAuthnCredentials(UserId, Nickname) WHERE Nickname IS NOT NULL;

-- FK → AspNetUsers(Id) : CASCADE 금지 + 동적 DROP은 변수 사용
DECLARE @fk sysname;

SELECT @fk = fk.name
FROM sys.foreign_keys fk
WHERE fk.parent_object_id   = OBJECT_ID(N'dbo.WebAuthnCredentials')
  AND fk.referenced_object_id = OBJECT_ID(N'dbo.AspNetUsers');

IF @fk IS NOT NULL
BEGIN
    DECLARE @stmt nvarchar(max) =
        N'ALTER TABLE dbo.WebAuthnCredentials DROP CONSTRAINT ' + QUOTENAME(@fk) + N';';
    EXEC(@stmt);
END;

ALTER TABLE dbo.WebAuthnCredentials
  ADD CONSTRAINT FK_WebAuthnCredentials_AspNetUsers_UserId
  FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id);  -- NO ACTION (기본)

-- 4) 해시 컬럼이 방금 추가되어 0x면 1회 백필
IF COL_LENGTH('dbo.WebAuthnCredentials','CredentialIdHash') IS NOT NULL
    UPDATE W SET CredentialIdHash = CASE WHEN CredentialIdHash = 0x THEN HASHBYTES('SHA2_256', W.CredentialId) ELSE CredentialIdHash END
    FROM dbo.WebAuthnCredentials W;
");


            migrationBuilder.CreateTable(
                name: "DepartmentMasterLoc",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DepartmentId = table.Column<int>(type: "int", nullable: false),
                    LangCode = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    Name = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: false),
                    ShortName = table.Column<string>(type: "nvarchar(32)", maxLength: 32, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DepartmentMasterLoc", x => x.Id);
                    table.ForeignKey(
                        name: "FK_DepartmentMasterLoc_DepartmentMasters_DepartmentId",
                        column: x => x.DepartmentId,
                        principalTable: "DepartmentMasters",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "PositionMasterLoc",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    PositionId = table.Column<int>(type: "int", nullable: false),
                    LangCode = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    Name = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: false),
                    ShortName = table.Column<string>(type: "nvarchar(32)", maxLength: 32, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PositionMasterLoc", x => x.Id);
                    table.ForeignKey(
                        name: "FK_PositionMasterLoc_PositionMasters_PositionId",
                        column: x => x.PositionId,
                        principalTable: "PositionMasters",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "UserProfiles",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    UserId = table.Column<string>(type: "nvarchar(450)", maxLength: 450, nullable: false),
                    CompCd = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    DisplayName = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: true),
                    DepartmentId = table.Column<int>(type: "int", nullable: true),
                    PositionId = table.Column<int>(type: "int", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_UserProfiles", x => x.Id);
                    table.ForeignKey(
                        name: "FK_UserProfiles_AspNetUsers_UserId",
                        column: x => x.UserId,
                        principalTable: "AspNetUsers",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                    table.ForeignKey(
                        name: "FK_UserProfiles_DepartmentMasters_DepartmentId",
                        column: x => x.DepartmentId,
                        principalTable: "DepartmentMasters",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.SetNull);
                    table.ForeignKey(
                        name: "FK_UserProfiles_PositionMasters_PositionId",
                        column: x => x.PositionId,
                        principalTable: "PositionMasters",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.SetNull);
                });

            migrationBuilder.CreateIndex(
                name: "IX_DepartmentMasterLoc_DepartmentId_LangCode",
                table: "DepartmentMasterLoc",
                columns: new[] { "DepartmentId", "LangCode" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_DepartmentMasters_CompCd_Code",
                table: "DepartmentMasters",
                columns: new[] { "CompCd", "Code" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_DepartmentMasters_CompCd_IsActive_SortOrder",
                table: "DepartmentMasters",
                columns: new[] { "CompCd", "IsActive", "SortOrder" });

            migrationBuilder.CreateIndex(
                name: "IX_PositionMasterLoc_PositionId_LangCode",
                table: "PositionMasterLoc",
                columns: new[] { "PositionId", "LangCode" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_PositionMasters_CompCd_Code",
                table: "PositionMasters",
                columns: new[] { "CompCd", "Code" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_PositionMasters_CompCd_IsActive_SortOrder",
                table: "PositionMasters",
                columns: new[] { "CompCd", "IsActive", "SortOrder" });

            migrationBuilder.CreateIndex(
                name: "IX_UserProfiles_DepartmentId",
                table: "UserProfiles",
                column: "DepartmentId");

            migrationBuilder.CreateIndex(
                name: "IX_UserProfiles_PositionId",
                table: "UserProfiles",
                column: "PositionId");

            migrationBuilder.CreateIndex(
                name: "IX_UserProfiles_UserId",
                table: "UserProfiles",
                column: "UserId",
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "DepartmentMasterLoc");

            migrationBuilder.DropTable(
                name: "PositionMasterLoc");

            migrationBuilder.DropTable(
                name: "UserProfiles");

            migrationBuilder.DropTable(
                name: "DepartmentMasters");

            migrationBuilder.DropTable(
                name: "PositionMasters");

            migrationBuilder.DropColumn(
                name: "Discriminator",
                table: "AspNetUsers");
        }
    }
}
