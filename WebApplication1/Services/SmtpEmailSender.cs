// 2025.12.11 Changed: 메일 본문 URL을 App PublicUrl 값으로 치환하고 제목 본문을 UTF8 인코딩으로 고정
using System.Net;
using System.Net.Mail;
using System.Text;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.Extensions.Configuration;
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

    public class SmtpEmailSender : IEmailSender
    {
        private readonly SmtpOptions _o;
        private readonly string? _publicBaseUrl;

        public SmtpEmailSender(IOptions<SmtpOptions> opt, IConfiguration cfg)
        {
            _o = opt.Value;
            _publicBaseUrl = cfg["App:PublicUrl"]?.TrimEnd('/');
        }

        public async Task SendEmailAsync(string email, string subject, string htmlMessage)
        {
            // 메일 본문 내 localhost 기반 URL을 App PublicUrl 값으로 치환
            if (!string.IsNullOrWhiteSpace(_publicBaseUrl) &&
                !string.IsNullOrEmpty(htmlMessage))
            {
                htmlMessage = htmlMessage.Replace("http://localhost:5000", _publicBaseUrl);
                htmlMessage = htmlMessage.Replace("https://localhost:5001", _publicBaseUrl);
            }

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
            };

            msg.To.Add(email);

            await client.SendMailAsync(msg);
        }
    }
}
