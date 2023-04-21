using Data_Server.Network;
using Event_Server.Network.ServerPacket;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;

namespace Event_Server.Network.ClientPacket
{
    public sealed class CpRequestLotteryInfo : IRecvPacket
    {
        private const byte MAX_BETS = 100;
        public void Process(byte[] buffer, IConnection connection)
        {
            var msg = new ByteBuffer(buffer);

            bool LotteryStatus = Conversoes.ByteToBoolean(msg.ReadByte());
            bool BetStatus = Conversoes.ByteToBoolean(msg.ReadByte());
            int Acumulado = msg.ReadInt32();
            byte LastBetNum = msg.ReadByte();
            string LastBetWinner = msg.ReadString();

            // Cria uma list, pra adicionar o nome e valor salvo nas apostas e posteriormente salvar em arquivo.
            List<(string, int)> nomeValor = new List<(string, int)>();
            for (int i= 1; i<= MAX_BETS; i++)
            {
                byte j = msg.ReadByte();
                if (j > 0)
                {
                    int valor = msg.ReadInt32();
                    string nome = msg.ReadString();

                    nomeValor.Add((nome, valor));
                }
            }

            // Abre ou cria um arquivo de texto chamado "arquivo.txt" na pasta atual
            using (StreamWriter writer = new StreamWriter("arquivo.txt"))
            {
                // Escreve uma linha de texto no arquivo
                writer.WriteLine("Olá, mundo!");
            }

            //Global.WriteLog(LogType.Player, mensagem, LogColor.Green);

            new SpLotteryData().Send(connection);
        }
    }
}