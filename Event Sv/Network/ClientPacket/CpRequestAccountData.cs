using Event_Server.Communication;
using Event_Server.Network.ServerPacket;
using Event_Server.Util;

namespace Event_Server.Network.ClientPacket {
    public sealed class CpRequestLotteryData : IRecvPacket  {
        public void Process(byte[] buffer, IConnection connection) {
            var msg = new ByteBuffer(buffer);

            var mensagem = msg.ReadString();

            Global.WriteLog(LogType.Player, mensagem, LogColor.Green);

            new SpAccountData().Send(connection);
            //SendPacket.Add(typeof(SpAccountData), (int)Packet.AccountData);
        }
    }
}