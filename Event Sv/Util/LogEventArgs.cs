using System;

namespace Event_Server.Util {
    public sealed class LogEventArgs : EventArgs {
        public string Text { get; private set; }
        public LogColor Color { get; private set; }
        public int Index { get; set; }

        public LogEventArgs(string text, LogColor color, int index) {
            Text = text;
            Color = color;
            Index = index;
        }
    }
}