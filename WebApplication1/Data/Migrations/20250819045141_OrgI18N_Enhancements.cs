using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    public partial class OrgI18N_Enhancements : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            // 기존 인덱스 제거
            migrationBuilder.DropIndex(
                name: "IX_PositionMasters_CompCd_IsActive_SortOrder",
                table: "PositionMasters");

            // RowVersion 추가
            migrationBuilder.AddColumn<byte[]>(
                name: "RowVersion",
                table: "UserProfiles",
                type: "rowversion",
                rowVersion: true,
                nullable: true);

            migrationBuilder.AddColumn<byte[]>(
                name: "RowVersion",
                table: "PositionMasters",
                type: "rowversion",
                rowVersion: true,
                nullable: true);

            migrationBuilder.AddColumn<byte[]>(
                name: "RowVersion",
                table: "PositionMasterLoc",
                type: "rowversion",
                rowVersion: true,
                nullable: true);

            migrationBuilder.AddColumn<byte[]>(
                name: "RowVersion",
                table: "DepartmentMasters",
                type: "rowversion",
                rowVersion: true,
                nullable: true);

            migrationBuilder.AddColumn<byte[]>(
                name: "RowVersion",
                table: "DepartmentMasterLoc",
                type: "rowversion",
                rowVersion: true,
                nullable: true);

            // 기본값 보강
            migrationBuilder.AlterColumn<int>(
                name: "SortOrder",
                table: "DepartmentMasters",
                type: "int",
                nullable: false,
                defaultValue: 0,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.AlterColumn<bool>(
                name: "IsActive",
                table: "DepartmentMasters",
                type: "bit",
                nullable: false,
                defaultValue: true,
                oldClrType: typeof(bool),
                oldType: "bit");

            // 신규 인덱스/체크 제약
            migrationBuilder.CreateIndex(
                name: "IX_PositionMasters_CompCd_IsActive_RankLevel_SortOrder",
                table: "PositionMasters",
                columns: new[] { "CompCd", "IsActive", "RankLevel", "SortOrder" });

            migrationBuilder.AddCheckConstraint(
                name: "CK_PositionMasters_RankLevel",
                table: "PositionMasters",
                sql: "[RankLevel] >= 0");

            migrationBuilder.CreateIndex(
                name: "IX_PositionMasterLoc_LangCode",
                table: "PositionMasterLoc",
                column: "LangCode");

            migrationBuilder.CreateIndex(
                name: "IX_DepartmentMasterLoc_LangCode",
                table: "DepartmentMasterLoc",
                column: "LangCode");

            migrationBuilder.AddCheckConstraint(
                name: "CK_DepartmentMasters_Code",
                table: "DepartmentMasters",
                sql: "LEN(LTRIM(RTRIM([Code]))) > 0");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            // 신규 인덱스/체크 제약 제거
            migrationBuilder.DropIndex(
                name: "IX_PositionMasters_CompCd_IsActive_RankLevel_SortOrder",
                table: "PositionMasters");

            migrationBuilder.DropCheckConstraint(
                name: "CK_PositionMasters_RankLevel",
                table: "PositionMasters");

            migrationBuilder.DropIndex(
                name: "IX_PositionMasterLoc_LangCode",
                table: "PositionMasterLoc");

            migrationBuilder.DropIndex(
                name: "IX_DepartmentMasterLoc_LangCode",
                table: "DepartmentMasterLoc");

            migrationBuilder.DropCheckConstraint(
                name: "CK_DepartmentMasters_Code",
                table: "DepartmentMasters");

            // RowVersion 컬럼 제거
            migrationBuilder.DropColumn(
                name: "RowVersion",
                table: "UserProfiles");

            migrationBuilder.DropColumn(
                name: "RowVersion",
                table: "PositionMasters");

            migrationBuilder.DropColumn(
                name: "RowVersion",
                table: "PositionMasterLoc");

            migrationBuilder.DropColumn(
                name: "RowVersion",
                table: "DepartmentMasters");

            migrationBuilder.DropColumn(
                name: "RowVersion",
                table: "DepartmentMasterLoc");

            // 기본값 원복
            migrationBuilder.AlterColumn<int>(
                name: "SortOrder",
                table: "DepartmentMasters",
                type: "int",
                nullable: false,
                oldClrType: typeof(int),
                oldType: "int",
                oldDefaultValue: 0);

            migrationBuilder.AlterColumn<bool>(
                name: "IsActive",
                table: "DepartmentMasters",
                type: "bit",
                nullable: false,
                oldClrType: typeof(bool),
                oldType: "bit",
                oldDefaultValue: true);

            // 예전 인덱스 복구
            migrationBuilder.CreateIndex(
                name: "IX_PositionMasters_CompCd_IsActive_SortOrder",
                table: "PositionMasters",
                columns: new[] { "CompCd", "IsActive", "SortOrder" });
        }
    }
}
