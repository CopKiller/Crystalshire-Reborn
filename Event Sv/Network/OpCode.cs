using System;
using System.Collections.Generic;
using Event_Server.Data;
using Event_Server.Network.ClientPacket;
using Event_Server.Network.ServerPacket;

namespace Event_Server.Network {
    public sealed class OpCode {
        public static Dictionary<int, Type> RecvPacket = new Dictionary<int, Type>();
        public static Dictionary<Type, int> SendPacket = new Dictionary<Type, int>();

        public static void InitOpCode() {
            // Fluxo
            // Servidor Principal  -->  Event Server
            // Recebendo dados Lottery
            RecvPacket.Add((int)Packet.LotteryData, typeof(CpReceiveLotteryData));
            // Recebendo dados Lottery
            RecvPacket.Add((int)Packet.LotteryInfo, typeof(CpRequestLotteryInfo));
            // Recebendo dados Items Pendentes
            RecvPacket.Add((int)Packet.ItemsPendentes, typeof(CpReceiveItemsPendentes));
            // Recebendo dados p/ DiscordBot
            RecvPacket.Add((int)Packet.DiscordMsg, typeof(CpReceiveDiscordMsg));
            // Recebendo dados p/ Servidor de Email
            RecvPacket.Add((int)Packet.AccountRecovery, typeof(CpReceiveAccountRecovery));

            // Fluxo
            // Event Server  -->  Servidor Principal
            // Enviando os dados salvo pro servidor principal
            SendPacket.Add(typeof(SpLotteryData), (int)Packet.LotteryData);
            // Enviando um ping, pra saber o status da conexão!
            SendPacket.Add(typeof(SpPing), (int)Packet.Ping);
            // Devolvendo retorno dos items pendentes!
            SendPacket.Add(typeof(SpItemsPendentes), (int)Packet.ItemsPendentes);

        }
    }
}