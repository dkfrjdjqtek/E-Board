using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplication1.Data.Migrations
{
    /// <inheritdoc />
    public partial class Add_IsAdmin_To_UserProfiles : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "IsAdmin",
                schema: "dbo",
                table: "UserProfiles",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "IsAdmin",
                schema: "dbo",
                table: "UserProfiles");
        }
    }
}
