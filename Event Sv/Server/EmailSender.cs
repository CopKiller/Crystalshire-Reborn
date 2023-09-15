using System;
using System.Collections.Generic;
using FluentEmail.Core;
using FluentEmail.Smtp;
using System.Net.Mail;
using System.Net;
using Event_Server.Communication;
using System.Threading.Tasks;
using System.Net.Http;

namespace Event_Server.Server
{
    enum EmailServer { Outlook = 1, Gmail = 2, Count }

    public class EmailServerInfo
    {
        public string ServerAddress { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
    }

    internal class EmailSender
    {
        // Pra salvar qual servidor de email foi usado.
        public EmailServer EmailServerUsed { get; private set; }

        // Referência dictionary com o endereço de cada servidor de email.
        public static Dictionary<EmailServer, EmailServerInfo> channelIds = new Dictionary<EmailServer, EmailServerInfo>()
        {
            { EmailServer.Outlook, new EmailServerInfo { ServerAddress = "smtp.office365.com",
                                                         Email = "email@hotmail.com",
                                                         Password = "123456" } },
            { EmailServer.Gmail, new EmailServerInfo { ServerAddress = "smtp.gmail.com",
                                                       Email = "xxx@gmail.com",
                                                       Password = "123456" } }
        };

        // Utilizado pra iniciar o componente global, método com o mesmo retorno do método global.
        public async Task<SmtpSender> InitEmailServerByServerIDAsync(EmailServer _serverId)
        {
            return await Task.Run(() =>
            {
                // Inicializa o componente a partir do ID do servidor.
                var smtpSender = new SmtpSender(() => new SmtpClient(
                    channelIds[_serverId].ServerAddress)
                {
                    Port = 587, EnableSsl = true, UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(
                    channelIds[_serverId].Email, channelIds[_serverId].Password)
                });

                EmailServerUsed = _serverId;

                return smtpSender;
            });
        }

        public static void SendEmail(string _from, string _to, string _subtitle, string _body, bool _isHtml = false)
        {
            Global.EmailSender.Send(Email.From(_from)
                             .To(_to)
                             .Subject(_subtitle)
                             .Body(_body, _isHtml));
        }
    }
}
