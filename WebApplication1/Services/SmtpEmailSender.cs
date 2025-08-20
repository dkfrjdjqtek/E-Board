using System.Net;
using System.Net.Mail;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.Extensions.Options;

namespace WebApplication1.Services // ← 프로젝트 이름에 맞추세요
{
    public class SmtpOptions
    {
        public string Host { get; set; } = "";
        public int Port { get; set; } = 587;
        public bool EnableSsl { get; set; } = true;
        public string User { get; set; } = "";
        public string Password { get; set; } = "";
        public string From { get; set; } = "";
        public string FromName { get; set; } = "E-Board";
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
                UseDefaultCredentials = false,                 // ← 꼭 false
                Timeout = 10000
            };

            var msg = new MailMessage
            {
                From = new MailAddress(_o.From, _o.FromName),
                Subject = subject,
                Body = htmlMessage,
                IsBodyHtml = true
            };
            msg.To.Add(email);

            await client.SendMailAsync(msg);
        }
    }
}
