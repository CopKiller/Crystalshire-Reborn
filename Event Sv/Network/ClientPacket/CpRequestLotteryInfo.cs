using Event_Server.Data;
using Event_Server.Network.ServerPacket;

namespace Event_Server.Network.ClientPacket
{
    public sealed class CpRequestLotteryInfo : IRecvPacket
    {
        public void Process(byte[] buffer, IConnection connection)
        {
            //var msg = new ByteBuffer(buffer);

            new SpLotteryData().SendPacket();
        }
    }
}