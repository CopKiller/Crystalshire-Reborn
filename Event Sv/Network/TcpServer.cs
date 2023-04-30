using System.Net;
using System.Linq;
using System.Net.Sockets;
using Event_Server.Util;
using Event_Server.Communication;
using System.Runtime.Remoting.Lifetime;

namespace Event_Server.Network {
    public sealed class TcpServer {
        public int Port { get; set; }

        private bool accept;
        private TcpListener server;

        public TcpServer() { }

        public TcpServer(int port) {
            Port = port;
        }

        public void InitServer() {
            Global.WriteLog(LogType.System, $"Initializing TCP Protocol...", LogColor.Coral);
            server = new TcpListener(IPAddress.Any, Port);
            server.Start();

            accept = true;
        }

        public void AcceptClient() {
            if (accept) {
                if (server.Pending()) {
                    var client = server.AcceptTcpClient();
                    var ipAddress = client.Client.RemoteEndPoint.ToString();

                    if (IsValidIpAddress(ipAddress)) {
                        var uniqueKey = new KeyGenerator().GetUniqueKey();

                        new Connection(client, ipAddress, uniqueKey);
                    }
                    else {
                        client.Close();
                        Global.WriteLog(LogType.System, $"Hacking Attempt: Invalid IpAddress {ipAddress}", LogColor.Blue);
                    }
                }
            }
        }

        public void Stop() {
            accept = false;
            server.Stop();
        }

        /// Envia um ping para determinar o estado da conexão.
        public void SendPing()
        {
            if (Connection.HighIndex > 0)
            {
                if (Connection.Connections != null && Connection.Connections.Count > 0)
                {
                    Connection.Connections[Connection.HighIndex].SendPing();
                }
                //ChangeState();
            }
        }
        /// Exibe a alteração no log quando o estado de conexão é alterado.
        //private void ChangeState()
        //{
        //    if (Connection.Connections[Connection.HighIndex].Connected != lastState)
        //    {
        //        if (Connection.Connections[Connection.HighIndex].Connected)
        //        {
        //            Global.WriteLog(LogType.System, "Main Server is connected", LogColor.Green);
        //        }
        //        else
        //        {
        //            Global.WriteLog(LogType.System, "Main Server is disconnected", LogColor.Red);
        //        }

        //        lastState = Connection.Connections[Connection.HighIndex].Connected;
        //    }
        //}

        private bool IsValidIpAddress(string ipAddress) {
            const int IpAddressArraySplit = 4;
            const int Last = 3;

            if (string.IsNullOrWhiteSpace(ipAddress) || string.IsNullOrEmpty(ipAddress)) {
                return false;
            }

            var values = ipAddress.Split('.');
            if (values.Length != IpAddressArraySplit) {
                return false;
            }

            // Retira o número da porta.
            values[Last] = values[Last].Remove(values[Last].IndexOf(':'));

            return values.All(r => byte.TryParse(r, out byte parsing));
        }
    }
}