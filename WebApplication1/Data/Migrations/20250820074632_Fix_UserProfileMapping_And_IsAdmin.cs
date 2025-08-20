using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    public partial class Fix_UserProfileMapping_And_IsAdmin : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            // ---- AspNetUsers 확장 컬럼 추가 ----
            migrationBuilder.AddColumn<int>(
                name: "DepartmentId",
                schema: "dbo",
                table: "AspNetUsers",
                type: "int",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "DisplayName",
                schema: "dbo",
                table: "AspNetUsers",
                type: "nvarchar(64)",
                maxLength: 64,
                nullable: true);

            migrationBuilder.AddColumn<int>(
                name: "IsAdmin",
                schema: "dbo",
                table: "AspNetUsers",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "PositionId",
                schema: "dbo",
                table: "AspNetUsers",
                type: "int",
                nullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_AspNetUsers_DepartmentId",
                schema: "dbo",
                table: "AspNetUsers",
                column: "DepartmentId");

            migrationBuilder.CreateIndex(
                name: "IX_AspNetUsers_PositionId",
                schema: "dbo",
                table: "AspNetUsers",
                column: "PositionId");

            migrationBuilder.AddForeignKey(
                name: "FK_AspNetUsers_DepartmentMasters_DepartmentId",
                schema: "dbo",
                table: "AspNetUsers",
                column: "DepartmentId",
                principalSchema: "dbo",
                principalTable: "DepartmentMasters",
                principalColumn: "Id");

            migrationBuilder.AddForeignKey(
                name: "FK_AspNetUsers_PositionMasters_PositionId",
                schema: "dbo",
                table: "AspNetUsers",
                column: "PositionId",
                principalSchema: "dbo",
                principalTable: "PositionMasters",
                principalColumn: "Id");

            // ---- UserProfiles의 PK를 UserId로 고정 (있으면 드롭 후 추가) ----
            migrationBuilder.Sql(@"
IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE name = N'PK_UserProfiles')
    ALTER TABLE [dbo].[UserProfiles] DROP CONSTRAINT [PK_UserProfiles];
IF COL_LENGTH('dbo.UserProfiles','UserId') IS NOT NULL
    ALTER TABLE [dbo].[UserProfiles] ADD CONSTRAINT [PK_UserProfiles] PRIMARY KEY ([UserId]);
");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            // FK/인덱스 제거
            migrationBuilder.DropForeignKey(
                name: "FK_AspNetUsers_DepartmentMasters_DepartmentId",
                schema: "dbo",
                table: "AspNetUsers");

            migrationBuilder.DropForeignKey(
                name: "FK_AspNetUsers_PositionMasters_PositionId",
                schema: "dbo",
                table: "AspNetUsers");

            migrationBuilder.DropIndex(
                name: "IX_AspNetUsers_DepartmentId",
                schema: "dbo",
                table: "AspNetUsers");

            migrationBuilder.DropIndex(
                name: "IX_AspNetUsers_PositionId",
                schema: "dbo",
                table: "AspNetUsers");

            // 확장 컬럼 제거
            migrationBuilder.DropColumn(
                name: "DepartmentId",
                schema: "dbo",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "DisplayName",
                schema: "dbo",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "IsAdmin",
                schema: "dbo",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "PositionId",
                schema: "dbo",
                table: "AspNetUsers");

            // UserProfiles PK 되돌림(안전 가드)
            migrationBuilder.Sql(@"
IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE name = N'PK_UserProfiles')
    ALTER TABLE [dbo].[UserProfiles] DROP CONSTRAINT [PK_UserProfiles];
IF COL_LENGTH('dbo.UserProfiles','Id') IS NOT NULL
    ALTER TABLE [dbo].[UserProfiles] ADD CONSTRAINT [PK_UserProfiles] PRIMARY KEY ([Id]);
");
        }
    }
}
