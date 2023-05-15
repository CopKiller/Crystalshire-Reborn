using Event_Server.Server;
using System;
using Event_Server.Communication;
using System.Windows.Forms;

namespace Event_Server {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            FrmMain frmMain = new FrmMain();
            frmMain.Show();

            Application.Idle += frmMain.OnApplicationIdle;
            Application.Run(frmMain);
        }
    }
}
