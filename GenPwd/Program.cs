// File: Tools/GenIdentityPasswordSql.cs
// 2025.09.18 Changed: IdentityUser 제거 → PasswordHasher<object> 사용(콘솔 환경 호환)
using System;
using Microsoft.AspNetCore.Identity;

class Program
{
    static void Main()
    {
        // === 필요값 수정 ===
        var userId = "8bf4629e-ea28-4ce5-b8e0-cbb74e65d579";   // 대상 사용자 Id
        var newPassword = "Admin!2345?";                       // 임시 비밀번호

        // 콘솔에서도 동작: TUser를 object로 대체
        var hasher = new PasswordHasher<object>();
        var hash = hasher.HashPassword(new object(), newPassword);

        var escapedHash = hash.Replace("'", "''"); // SQL 이스케이프
        Console.WriteLine("-- Paste this into SSMS on DB: LCHYEBOARD");
        Console.WriteLine("UPDATE LCHYEBOARD.dbo.AspNetUsers");
        Console.WriteLine($"SET PasswordHash = '{escapedHash}',");
        Console.WriteLine("    SecurityStamp = NEWID(),");
        Console.WriteLine("    AccessFailedCount = 0,");
        Console.WriteLine("    LockoutEnd = NULL");
        Console.WriteLine($"WHERE Id = '{userId}';");
        Console.WriteLine();
        Console.WriteLine("-- 로그인 비번: " + newPassword + " (즉시 변경 권장)");
    }
}
