using System;
using System.Windows.Forms;
using System.Drawing;
using System.Security;
using System.Threading;
using System.Runtime.InteropServices;
using Event_Server.Util;
using Event_Server.Server;
using Event_Server.Communication;
using Event_Server.Network.ServerPacket;
using Event_Server.Data;
using Event_Server.Network;
using System.Threading.Tasks;
using FluentEmail.Smtp;
using FluentEmail.Core;
using static System.Net.WebRequestMethods;

namespace Event_Server
{
    public partial class FrmMain : Form
    {

        #region Peek Message
        [SuppressUnmanagedCodeSecurity]
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern bool PeekMessage(out Message msg, IntPtr hWnd, uint messageFilterMin, uint messageFilterMax, uint flags);

        [StructLayout(LayoutKind.Sequential)]
        private struct Message
        {
            public IntPtr hWnd;
            public IntPtr msg;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public Point p;
        }
        public void OnApplicationIdle(object sender, EventArgs e)
        {
            while (this.AppStillIdle)
            {
                if (Server != null)
                {
                    Server.ServerLoop();

                    Thread.Sleep(1);

                    if (!Server.ServerRunning)
                    {
                        Server.StopServer();
                        Environment.Exit(0);
                    }
                }
            }
        }

        private bool AppStillIdle
        {
            get
            {
                return !PeekMessage(out Message msg, IntPtr.Zero, 0, 0, 0);
            }
        }

        #endregion

        DataServer Server;

        const int MaxLogsLines = 250;
        const int PreserveLogsLines = 25;

        enum CloseSteps
        {
            None,
            Close,
        }

        public FrmMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            InitServerAsync();
        }

        private async Task InitServerAsync()
        {
            await InitLogs();
            await InitDiscordBotAsync();
            await InitEmailSenderAsync();
            Server = new DataServer();
            Server.UpdateUps += UpdateUps;
            Server.InitServer();
        }

        private async Task InitEmailSenderAsync()
        {
            //Inicializa por padrão no servidor do outlook.
            var _emailServer = EmailServer.Outlook;

            Global.SystemLogs.Write($"Initializing Email Server SMTP({Enum.GetName(typeof(EmailServer), _emailServer)})...", LogColor.Coral);
            Global.EmailSender = await new EmailSender().InitEmailServerByServerIDAsync(_emailServer);

        }

        private async Task InitDiscordBotAsync()
        {
            Global.SystemLogs.Write("Initializing DiscordBot...", LogColor.Coral);
            Global.DiscordBot = new DiscordBot();
            await Global.DiscordBot.Start();
        }

        private Task InitLogs()
        {
            //System
            Global.SystemLogs = new Log("System")
            {
                Index = 0
            };

            Global.SystemLogs.LogEvent += WriteLog;

            var result = Global.SystemLogs.OpenFile();

            if (result)
            {
                Global.SystemLogs.Enabled = true;
            }
            else
            {
                MessageBox.Show("An error ocurred when trying to open the file log.");
            }

            //Player
            Global.PlayerLogs = new Log("Player")
            {
                Index = 1
            };

            Global.PlayerLogs.LogEvent += WriteLog;

            result = Global.PlayerLogs.OpenFile();

            if (result)
            {
                Global.PlayerLogs.Enabled = true;
            }
            else
            {
                MessageBox.Show("An error ocurred when trying to open the file log.");
            }

            //Debug
            Global.DebugLogs = new Log("Debug")
            {
                Index = 2
            };

            Global.DebugLogs.LogEvent += WriteLog;

            result = Global.DebugLogs.OpenFile();

            if (result)
            {
                Global.DebugLogs.Enabled = true;
            }
            else
            {
                MessageBox.Show("An error ocurred when trying to open the file log.");
            }

            Global.WriteLog(LogType.System, $"Initializing Logs...", LogColor.Coral);
            return Task.CompletedTask;
        }
        private void WriteLog(object sender, LogEventArgs e)
        {
            var text = TextSystem;

            switch ((LogType)e.Index)
            {
                case LogType.System:
                    text = TextSystem;
                    break;
            }

            text.SelectionStart = text.TextLength;
            text.SelectionLength = 0;
            text.SelectionColor = GetColor(e.Color);
            text.AppendText($"{DateTime.Now}: {e.Text}{Environment.NewLine}");
            text.ScrollToCaret();
        }

        private void UpdateUps(int ups)
        {
            Text = $"Event Server @ {ups} Ups";
        }

        private void MenuExit_Click(object sender, EventArgs e)
        {
            CheckForCloseApplication();
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            CheckForCloseApplication();
        }

        public void AddText(string text)
        {
            Text = text;
        }

        private Color GetColor(LogColor logColor)
        {
            Color color = Color.Empty;

            switch (logColor)
            {
                case LogColor.Black:
                    color = Color.Black;
                    break;
                case LogColor.Blue:
                    color = Color.Black;
                    break;
                case LogColor.Coral:
                    color = Color.Coral;
                    break;
                case LogColor.Green:
                    color = Color.Green;
                    break;
                case LogColor.Red:
                    color = Color.Red;
                    break;
            }

            return color;
        }

        private void CheckForCloseApplication()
        {
            var steps = GetCloseApplicationStep();
            if (steps == CloseSteps.Close)
            {
                Server.ServerRunning = false;
            }
        }

        private CloseSteps GetCloseApplicationStep()
        {
            var closeSteps = CloseSteps.None;
            var msg = "Do you want to exit?";
            var caption = "Question";

            var result = MessageBox.Show(msg, caption, MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                closeSteps = CloseSteps.Close;
            }

            return closeSteps;
        }

        private void button1_Click(object sender, EventArgs e, SpLotteryData packet)
        {
            //var Conexao = new Connection();
            //var SPacket = new SpAccountData();

            //new SpAccountData().Send(connection);

            //var Packet = new ByteBuffer();
            //Packet.Write("Toma porra Loka");

            //SPacket.Send(Conexao);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if (Connection.HighIndex > 0)
            //{
            //    new SpLotteryData(new Lottery().Load()).Send(Connection.Connections[Connection.HighIndex]);
            //}

            EmailSender.SendEmail(_from: "felipe_157@windowslive.com",
                                  _to: "feliposos22@gmail.com",
                                  _subtitle: "Exemplo de e-mail com FluentEmail",
                                  _body: "Testandooo");

                                 //Exemplo Html:  _body: "< !DOCTYPE html >< html >< head >< meta http - equiv =\"Content-Type\" content=\"text/html; charset=UTF-8\"><link rel=\"stylesheet\" type=\"text/css\" id=\"u0\" href=\"https://pt.rakko.tools/tools/129/lib/tinymce/skins/ui/oxide/content.min.css\"><link rel=\"stylesheet\" type=\"text/css\" id=\"u1\" href=\"https://pt.rakko.tools/tools/129/lib/tinymce/skins/content/default/content.min.css\"></head><body id=\"tinymce\" class=\"mce-content-body \" data-id=\"content\" contenteditable=\"true\" spellcheck=\"false\"><p><img src=\"https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExMTQwNjk1NjI2NzZjMDcxMTBkMmIwOGE5N2NlZjY3ZWNmOWE5MDJjYiZlcD12MV9pbnRlcm5hbF9naWZzX2dpZklkJmN0PXM/wh5frFkevniZXJpYgd/giphy.gif\" data-mce-src=\"https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExMTQwNjk1NjI2NzZjMDcxMTBkMmIwOGE5N2NlZjY3ZWNmOWE5MDJjYiZlcD12MV9pbnRlcm5hbF9naWZzX2dpZklkJmN0PXM/wh5frFkevniZXJpYgd/giphy.gif\"> Olá, este é um email utilizando HTML.</p></body></html>", true);
        }
    }
}