using Event_Server.Data;
using Event_Server.Network.ServerPacket;
using System.Runtime.InteropServices;

namespace Event_Server.Network.ClientPacket
{
    public sealed class CpReceiveAccountRecovery : IRecvPacket
    {
        public void Process(byte[] buffer, IConnection connection)
        {
            var msg = new ByteBuffer(buffer);

            //new CpReceiveRecoveryAccount(new Lottery().Load()).Send(connection);

            var email = msg.ReadString();

            var password = msg.ReadString();

        }
    }
}