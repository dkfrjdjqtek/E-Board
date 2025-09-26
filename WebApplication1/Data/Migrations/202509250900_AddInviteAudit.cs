using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

// 2025.09.25 Added: InviteAudits 테이블 신규 생성 PK만 생성 FK 생성 안함
public partial class AddInviteAudit : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        // 2025.09.25 Added: 초대 메일 발송 이력 테이블 생성
        migrationBuilder.CreateTable(
            name: "InviteAudits",
            columns: table => new
            {
                Id = table.Column<int>(type: "int", nullable: false)
                    .Annotation("SqlServer:Identity", "1, 1"), // PK
                UserId = table.Column<string>(type: "nvarchar(450)", nullable: false),
                Email = table.Column<string>(type: "nvarchar(256)", maxLength: 256, nullable: false),
                SentAtUtc = table.Column<DateTime>(type: "datetime2", nullable: false),
                Status = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: false),
                ErrorMessage = table.Column<string>(type: "nvarchar(max)", nullable: true),
                CorrelationId = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true)
            },
            constraints: table =>
            {
                table.PrimaryKey("PK_InviteAudits", x => x.Id); // PK만 생성
            });

        // 2025.09.25 Added: 조회 성능 향상을 위한 보조 인덱스 두 개 추가 UK 아님
        migrationBuilder.CreateIndex(
            name: "IX_InviteAudits_UserId_SentAtUtc",
            table: "InviteAudits",
            columns: new[] { "UserId", "SentAtUtc" });

        migrationBuilder.CreateIndex(
            name: "IX_InviteAudits_Email_SentAtUtc",
            table: "InviteAudits",
            columns: new[] { "Email", "SentAtUtc" });
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
        // 2025.09.25 Added: 롤백 시 테이블 삭제
        migrationBuilder.DropTable(
            name: "InviteAudits");
    }
}
