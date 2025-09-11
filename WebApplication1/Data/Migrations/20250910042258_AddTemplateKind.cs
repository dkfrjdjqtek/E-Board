using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    /// <inheritdoc />
    public partial class AddTemplateKind : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.EnsureSchema(
                name: "dbo");

            // ===== dbo.TemplateKindMasters =====
            migrationBuilder.CreateTable(
                name: "TemplateKindMasters",
                schema: "dbo",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                              .Annotation("SqlServer:Identity", "1, 1"),
                    CompCd = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    DepartmentId = table.Column<int>(type: "int", nullable: false, defaultValue: 0),
                    Code = table.Column<string>(type: "nvarchar(32)", maxLength: 32, nullable: false),
                    Name = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: false),
                    IsActive = table.Column<bool>(type: "bit", nullable: false, defaultValue: true),
                    RowVersion = table.Column<byte[]>(type: "rowversion", rowVersion: true, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_TemplateKindMasters", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_TemplateKindMasters_CompCd_Code",
                schema: "dbo",
                table: "TemplateKindMasters",
                columns: new[] { "CompCd", "Code" },
                unique: true);

            migrationBuilder.AddCheckConstraint(
                name: "CK_TemplateKindMasters_Code",
                schema: "dbo",
                table: "TemplateKindMasters",
                sql: "LEN(LTRIM(RTRIM([Code]))) > 0");

            // ===== dbo.TemplateKindMasterLoc =====
            migrationBuilder.CreateTable(
                name: "TemplateKindMasterLoc",
                schema: "dbo",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false),
                    LangCode = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    Name = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: false),
                    RowVersion = table.Column<byte[]>(type: "rowversion", rowVersion: true, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_TemplateKindMasterLoc", x => new { x.Id, x.LangCode });
                    table.ForeignKey(
                        name: "FK_TemplateKindMasterLoc_TemplateKindMasters_Id",
                        column: x => x.Id,
                        principalSchema: "dbo",
                        principalTable: "TemplateKindMasters",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_TemplateKindMasterLoc_LangCode",
                schema: "dbo",
                table: "TemplateKindMasterLoc",
                column: "LangCode");

            migrationBuilder.AddCheckConstraint(
                name: "CK_TemplateKindMasterLoc_Name",
                schema: "dbo",
                table: "TemplateKindMasterLoc",
                sql: "LEN(LTRIM(RTRIM([Name]))) > 0");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            // FK 의존성 때문에 Loc 먼저 Drop
            migrationBuilder.DropTable(
                name: "TemplateKindMasterLoc",
                schema: "dbo");

            migrationBuilder.DropTable(
                name: "TemplateKindMasters",
                schema: "dbo");
        }
    }
}
