namespace WebApplication1.Services
{
    public interface IAuditLogger
    {
        Task LogAsync(string docId, string actorId, string actionCode, string? detailJson);
    }
}