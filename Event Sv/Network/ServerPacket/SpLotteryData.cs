using System.Collections.Generic;
using Data_Server.Network;
using Event_Server.Communication;
using Event_Server.Data;

namespace Event_Server.Network.ServerPacket
{
    public sealed class SpLotteryData : SendPacket
    {
        public SpLotteryData(Lottery LoadLottery)
        {

            msg = new ByteBuffer();
            msg.Write(OpCode.SendPacket[GetType()]);
            msg.Write(Conversoes.BooleanToByte(LoadLottery.LotteryStatus));
            msg.Write(Conversoes.BooleanToByte(LoadLottery.BetStatus));
            msg.Write(LoadLottery.Acumulado);
            msg.Write(LoadLottery.LastBetNum);
            msg.Write(LoadLottery.LastBetWinner.Trim());

            if (LoadLottery.Apostas.Count > 0)
            {
                //Envia a quantidade de items do array
                msg.Write(LoadLottery.Apostas.Count);

                for (int i = 0; i < LoadLottery.Apostas.Count; i++)
                {
                    //Numero Apostado
                    msg.Write(LoadLottery.Apostas[i].Item1);
                    //Nome
                    msg.Write(LoadLottery.Apostas[i].Item2.Trim());
                    //Valor
                    msg.Write(LoadLottery.Apostas[i].Item3);
                }
            }
        }

        // Exemplo de como enviar uma packet pro client sem precisar do servidor principal acioná-la!
        //public void SendPacket(Lottery LoadLottery)
        //{
        //    if (Connection.HighIndex > 0)
        //    {
        //        new SpLotteryData(LoadLottery).Send(Connection.Connections[Connection.HighIndex]);
        //    }
        //}
    }
}