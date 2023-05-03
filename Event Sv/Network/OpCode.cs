using System;
using System.Collections.Generic;
using Event_Server.Network.ClientPacket;
using Event_Server.Network.ServerPacket;

namespace Event_Server.Network {
    public sealed class OpCode {
        public static Dictionary<int, Type> RecvPacket = new Dictionary<int, Type>();
        public static Dictionary<Type, int> SendPacket = new Dictionary<Type, int>();

        public static void InitOpCode() {
            // Servidor Principal  -->  Event Server
            // Recebendo dados a serem salvos!
            RecvPacket.Add((int)Packet.LotteryData, typeof(CpReceiveLotteryData));
            // Servidor solicitando o envio dos dados da loteria!
            RecvPacket.Add((int)Packet.LotteryInfo, typeof(CpRequestLotteryInfo));

            // Event Server  -->  Servidor Principal
            // Enviando os dados salvo pro servidor principal
            SendPacket.Add(typeof(SpLotteryData), (int)Packet.LotteryData);
            // Enviando um ping, pra saber o status da conexão!
            SendPacket.Add(typeof(SpPing), (int)Packet.Ping);

        }
    }
}