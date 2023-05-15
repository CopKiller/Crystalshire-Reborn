using Event_Server.Data;
using Event_Server.Network.ServerPacket;
using System.Runtime.InteropServices;

namespace Event_Server.Network.ClientPacket
{
    public sealed class CpRequestLotteryInfo : IRecvPacket
    {
        public void Process(byte[] buffer, IConnection connection)
        {
            new SpLotteryData(new Lottery().Load()).Send(connection);

        }
    }
}