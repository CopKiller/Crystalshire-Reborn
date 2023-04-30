using System;
using Event_Server.Network;
using Event_Server.Communication;
using Event_Server.Util;
//using Event_Server.Cryptography;


namespace Event_Server.Server {
    public class DataServer {
        public Action<int> UpdateUps;
        public bool ServerRunning { get; set; } = true;

        private int tick;
        private int count;
        private int ups;

        TcpServer Server;

        public void InitServer() {
            //ConnectionPassword.InitializePassword();
            Server = new TcpServer(Constants.PORT);
            Server.InitServer();

            OpCode.InitOpCode();
        }

        public void ServerLoop() {
            Server.AcceptClient();

            Server.SendPing();
            ReceiveSocketData();

            CountUps();
        }

        public void StopServer() {
            Server.Stop();
            Connection.Connections.Clear();
        }

        private void CountUps() {
            if (Environment.TickCount >= tick + 1000) {
                ups = count;
                count = 0;
                tick = Environment.TickCount;

                UpdateUps?.Invoke(ups);
            }
            else {
                count++;
            }
        }

        private void ReceiveSocketData() {
            if (Connection.HighIndex == 0) { return; }
                Connection.Connections[Connection.HighIndex].ReceiveData();

                RemoveWhenNotConnected(Connection.HighIndex);
        }

        private void RemoveWhenNotConnected(int index) {
            if (!Connection.Connections[index].Connected) {
                string ipAddress = Connection.Connections[index].IpAddress;
                string uniqueKey = Connection.Connections[index].UniqueKey;

                //Connection.Connections[index].
                Connection.Remove(index);

                Global.WriteLog(LogType.System, $"{ipAddress} Key {uniqueKey} is disconnected", LogColor.Coral);
            }
        }
    }
}