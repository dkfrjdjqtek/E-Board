using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    /// <inheritdoc />
    public partial class Sync_Context_To_Db : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_DepartmentMasterLoc_DepartmentMasters_DepartmentId",
                table: "DepartmentMasterLoc");

            migrationBuilder.DropForeignKey(
                name: "FK_PositionMasterLoc_PositionMasters_PositionId",
                table: "PositionMasterLoc");

            migrationBuilder.DropPrimaryKey(
                name: "PK_PositionMasterLoc",
                table: "PositionMasterLoc");

            migrationBuilder.DropPrimaryKey(
                name: "PK_DepartmentMasterLoc",
                table: "DepartmentMasterLoc");

            migrationBuilder.DropColumn(
                name: "Discriminator",
                table: "AspNetUsers");

            migrationBuilder.RenameTable(
                name: "PositionMasterLoc",
                newName: "PositionMasterLocs");

            migrationBuilder.RenameTable(
                name: "DepartmentMasterLoc",
                newName: "DepartmentMasterLocs");

            migrationBuilder.RenameIndex(
                name: "IX_PositionMasterLoc_PositionId_LangCode",
                table: "PositionMasterLocs",
                newName: "IX_PositionMasterLocs_PositionId_LangCode");

            migrationBuilder.RenameIndex(
                name: "IX_DepartmentMasterLoc_DepartmentId_LangCode",
                table: "DepartmentMasterLocs",
                newName: "IX_DepartmentMasterLocs_DepartmentId_LangCode");

            migrationBuilder.AddPrimaryKey(
                name: "PK_PositionMasterLocs",
                table: "PositionMasterLocs",
                column: "Id");

            migrationBuilder.AddPrimaryKey(
                name: "PK_DepartmentMasterLocs",
                table: "DepartmentMasterLocs",
                column: "Id");

            migrationBuilder.AddForeignKey(
                name: "FK_DepartmentMasterLocs_DepartmentMasters_DepartmentId",
                table: "DepartmentMasterLocs",
                column: "DepartmentId",
                principalTable: "DepartmentMasters",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_PositionMasterLocs_PositionMasters_PositionId",
                table: "PositionMasterLocs",
                column: "PositionId",
                principalTable: "PositionMasters",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_DepartmentMasterLocs_DepartmentMasters_DepartmentId",
                table: "DepartmentMasterLocs");

            migrationBuilder.DropForeignKey(
                name: "FK_PositionMasterLocs_PositionMasters_PositionId",
                table: "PositionMasterLocs");

            migrationBuilder.DropPrimaryKey(
                name: "PK_PositionMasterLocs",
                table: "PositionMasterLocs");

            migrationBuilder.DropPrimaryKey(
                name: "PK_DepartmentMasterLocs",
                table: "DepartmentMasterLocs");

            migrationBuilder.RenameTable(
                name: "PositionMasterLocs",
                newName: "PositionMasterLoc");

            migrationBuilder.RenameTable(
                name: "DepartmentMasterLocs",
                newName: "DepartmentMasterLoc");

            migrationBuilder.RenameIndex(
                name: "IX_PositionMasterLocs_PositionId_LangCode",
                table: "PositionMasterLoc",
                newName: "IX_PositionMasterLoc_PositionId_LangCode");

            migrationBuilder.RenameIndex(
                name: "IX_DepartmentMasterLocs_DepartmentId_LangCode",
                table: "DepartmentMasterLoc",
                newName: "IX_DepartmentMasterLoc_DepartmentId_LangCode");

            migrationBuilder.AddColumn<string>(
                name: "Discriminator",
                table: "AspNetUsers",
                type: "nvarchar(21)",
                maxLength: 21,
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddPrimaryKey(
                name: "PK_PositionMasterLoc",
                table: "PositionMasterLoc",
                column: "Id");

            migrationBuilder.AddPrimaryKey(
                name: "PK_DepartmentMasterLoc",
                table: "DepartmentMasterLoc",
                column: "Id");

            migrationBuilder.AddForeignKey(
                name: "FK_DepartmentMasterLoc_DepartmentMasters_DepartmentId",
                table: "DepartmentMasterLoc",
                column: "DepartmentId",
                principalTable: "DepartmentMasters",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_PositionMasterLoc_PositionMasters_PositionId",
                table: "PositionMasterLoc",
                column: "PositionId",
                principalTable: "PositionMasters",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
