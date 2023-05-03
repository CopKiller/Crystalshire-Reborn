using Data_Server.Network;
using Event_Server.Communication;
using Event_Server.Data;
using Event_Server.Network.ServerPacket;
using Event_Server.Util;
using System.Collections.Generic;

namespace Event_Server.Network.ClientPacket
{
    public sealed class CpReceiveLotteryData : IRecvPacket  {
        public void Process(byte[] buffer, IConnection connection) {
            var msg = new ByteBuffer(buffer);

            var lotteryData = new Lottery();

            lotteryData.LotteryStatus = Conversoes.ByteToBoolean(msg.ReadByte());
            lotteryData.BetStatus = Conversoes.ByteToBoolean(msg.ReadByte());
            lotteryData.Acumulado = msg.ReadInt32();
            lotteryData.LastBetNum = msg.ReadByte();
            lotteryData.LastBetWinner = msg.ReadString().Trim();

            // Cria uma lista com as apostas.
            var apostas = new List<(byte, string, int)>();
            for (int i = 1; i <= Constants.MAX_BETS; i++)
            {
                byte j = msg.ReadByte();
                if (j > 0)
                {
                    int valor = msg.ReadInt32();
                    string nome = msg.ReadString().Trim();

                    apostas.Add((j, nome, valor));
                }
            }
            lotteryData.Apostas = apostas;

            lotteryData.Save(lotteryData);

            Global.WriteLog(LogType.Player, "Lottery Data Received and Saved!", LogColor.Green);

            //new SpLotteryData().Send(connection);
        }
    }
}