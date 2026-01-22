using System.Threading.Tasks;

namespace WebApplication1.Services
{
    public interface IWebPushNotifier
    {
        Task<bool> SendToUserIdAsync(string userId, string title, string body, string url, string? tag = null);
    }
}