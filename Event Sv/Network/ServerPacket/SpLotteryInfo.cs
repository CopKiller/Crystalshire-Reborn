using System.Collections.Generic;
using Event_Server.Data;

namespace Event_Server.Network.ServerPacket {
    public sealed class SpLotteryInfo : SendPacket {
        public SpLotteryInfo() {

            byte one = 25;
            int one1 = 99999999;
            msg = new ByteBuffer();
            msg.Write(OpCode.SendPacket[GetType()]);
            msg.Write("Tomaa k7");
            msg.Write(one);
            msg.Write(one1);
        }
    }
}