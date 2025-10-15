// 2025.10.14 Added: SQL Server 직접 INSERT로 dbo.DocumentAuditLogs 적재
// 2025.10.14 Added: IConfiguration 기반 연결 문자열 사용 ("DefaultConnection" 가정)
// 주의: 테이블명은 요청에 맞춰 DocumentAuditLogs 로 반영
using System.Data;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace WebApplication1.Services
{
    public class AuditLoggerSql : IAuditLogger
    {
        private readonly string _connStr;

        public AuditLoggerSql(IConfiguration cfg)
        {
            _connStr = cfg.GetConnectionString("DefaultConnection") ?? throw new InvalidOperationException("Missing connection string: DefaultConnection");
        }

        public async Task LogAsync(string docId, string actorId, string actionCode, string? detailJson)
        {
            // 2025.10.14 Added: INSERT 문 - PK는 IDENTITY, FK 없음, UTC 시간은 DB 기본값 사용 가능
            const string SQL = @"
INSERT INTO dbo.DocumentAuditLogs (DocId, ActorId, ActionCode, Detail)
VALUES (@DocId, @ActorId, @ActionCode, @Detail);";

            using var conn = new SqlConnection(_connStr);
            using var cmd = new SqlCommand(SQL, conn) { CommandType = CommandType.Text };
            cmd.Parameters.AddWithValue("@DocId", docId ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@ActorId", actorId ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@ActionCode", actionCode ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@Detail", (object?)detailJson ?? DBNull.Value);

            await conn.OpenAsync().ConfigureAwait(false);
            await cmd.ExecuteNonQueryAsync().ConfigureAwait(false);
        }
    }
}
