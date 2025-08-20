using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    /// <inheritdoc />
    public partial class Add_CompMaster_And_UserProfile_Comp_Default : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "CompMasters",
                schema: "dbo",
                columns: table => new
                {
                    CompCd = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    Name = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: false),
                    IsActive = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_CompMasters", x => x.CompCd);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "CompMasters",
                schema: "dbo");
        }
    }
}
