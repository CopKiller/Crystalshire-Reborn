using Event_Server.Network.ServerPacket;
using System.Collections.Generic;
using System.Text.Json;
using System.IO;
using Event_Server.Util;
using Data_Server.Util;
using Event_Server.Network;
using Event_Server.Communication;
using System.Net;

namespace Event_Server.Data
{

    public sealed class Lottery
    {
        // Directory
        private const string DirLottery = @"~\Lottery";

        public bool LotteryStatus { get; set; }
        public bool BetStatus { get; set; }
        public int Acumulado { get; set; }
        public byte LastBetNum { get; set; }
        public string LastBetWinner { get; set; }
        public List<(byte, string, int)> Apostas { get; set; }

        private string ArchiveName = @"/lotteryData.json";

        public Lottery() { }
        public void Save(Lottery lottery)
        {
            // Cria o diretório se não existir
            if (!Directory.Exists(DirLottery.MyDir()))
            {
                Directory.CreateDirectory(DirLottery.MyDir());
            }

            // Adiciona o conversor de tuplas
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new TupleConverterBSI() }
            };

            string json = JsonSerializer.Serialize(lottery, options);
            File.WriteAllText(DirLottery.MyDir() + ArchiveName, json);
        }
        public Lottery Load()
        {
            string filePath = DirLottery.MyDir() + ArchiveName;
            if (!File.Exists(filePath)) return null;

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new TupleConverterBSI() }
            };

            string json = File.ReadAllText(filePath);

            var ResultJson = JsonSerializer.Deserialize<Lottery>(json, options);

            return ResultJson;
        }
    }
}