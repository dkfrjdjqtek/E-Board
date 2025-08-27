using System.Net;
using System.Net.Mail;
using System.Text;                    // ← 추가
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.Extensions.Options;

namespace WebApplication1.Services
{
    public class SmtpOptions
    {
        public string Host { get; set; } = "";
        public int Port { get; set; } = 587;
        public bool EnableSsl { get; set; } = true;
        public string User { get; set; } = "";
        public string Password { get; set; } = "";
        public string From { get; set; } = "";
        public string FromName { get; set; } = "Han Young E-Board";
    }

    public class SmtpEmailSender(IOptions<SmtpOptions> opt) : IEmailSender
    {
        private readonly SmtpOptions _o = opt.Value;

        public async Task SendEmailAsync(string email, string subject, string htmlMessage)
        {
            using var client = new SmtpClient(_o.Host, _o.Port)
            {
                EnableSsl = _o.EnableSsl,
                Credentials = new NetworkCredential(_o.User, _o.Password),
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Timeout = 10000
            };

            using var msg = new MailMessage
            {
                From = new MailAddress(_o.From, _o.FromName),
                Subject = subject,
                SubjectEncoding = Encoding.UTF8,   // 제목 인코딩
                Body = htmlMessage,
                BodyEncoding = Encoding.UTF8,      // 본문 인코딩
                IsBodyHtml = true
                // HeadersEncoding = Encoding.UTF8  // (선택) 헤더 인코딩까지 강제하고 싶다면
            };
            msg.To.Add(email);

            await client.SendMailAsync(msg);
        }
    }
}
