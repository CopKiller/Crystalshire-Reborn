using Event_Server.Cryptography;
using Event_Server.Server;
using Event_Server.Util;
using FluentEmail.Core;
using FluentEmail.Smtp;
using System.Net;
using System.Net.Mail;

namespace Event_Server.Communication
{
    public static class Global
    {
        public static int Environment;
        public static Log PlayerLogs { get; set; }
        public static Log SystemLogs { get; set; }
        public static Log DebugLogs { get; set; }
        public static DiscordBot DiscordBot { get; set; }
        public static SmtpSender EmailSender { get; set; }

        public static void WriteLog(LogType type, string text, LogColor color)
        {
            switch (type)
            {
                case LogType.Player:
                    PlayerLogs.Write(text, color);
                    break;
                case LogType.System:
                    SystemLogs.Write(text, color);
                    break;
                case LogType.Debug: 
                    DebugLogs.Write(text, color); 
                    break;
            }
        }
    }
}