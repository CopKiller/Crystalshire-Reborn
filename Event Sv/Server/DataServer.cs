using System;
using Event_Server.Network;
using Event_Server.Communication;

namespace Event_Server.Server {
    public class DataServer {
        public Action<int> UpdateUps;
        public bool ServerRunning { get; set; } = true;

        private int tick;
        private int count;
        private int ups;

        TcpServer Server;
    
        public void InitServer() {
            Server = new TcpServer(Constants.PORT);
            Server.InitServer();

            OpCode.InitOpCode();
        }

        public void ServerLoop() {
            Server.AcceptClient();

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
            for (int n = 1; n <= Connection.HighIndex; n++) {
                if (Connection.Connections.ContainsKey(n)) {
                    Connection.Connections[n].ReceiveData();

                    RemoveWhenNotConnected(n);
                }
            }
        }

        private void RemoveWhenNotConnected(int index) {
            if (!Connection.Connections[index].Connected) {
                string ipAddress = Connection.Connections[index].IpAddress;
                string uniqueKey = Connection.Connections[index].UniqueKey;

                Connection.Remove(index);

                //WriteLog(LogType.System, $"{ipAddress} Key {uniqueKey} is disconnected", LogColor.Coral);
            }
        }
    }
}