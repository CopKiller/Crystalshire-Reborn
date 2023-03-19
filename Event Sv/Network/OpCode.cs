using System;
using System.Collections.Generic;
using Event_Server.Network.ClientPacket;
using Event_Server.Network.ServerPacket;

namespace Event_Server.Network {
    public sealed class OpCode {
        public static Dictionary<int, Type> RecvPacket = new Dictionary<int, Type>();
        public static Dictionary<Type, int> SendPacket = new Dictionary<Type, int>();

        public static void InitOpCode() {
            RecvPacket.Add((int)Packet.LotteryData, typeof(CpRequestLotteryData));

            SendPacket.Add(typeof(SpAccountData), (int)Packet.LotteryData);
        }
    }
}