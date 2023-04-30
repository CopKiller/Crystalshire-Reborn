using Event_Server.Util;
using System;

namespace Event_Server.Communication
{
    public class Global
    {
        public int Environment;
        public static Log PlayerLogs { get; set; }
        public static Log SystemLogs { get; set; }
        public static Log DebugLogs { get; set; }

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