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
        public bool LotteryStatus { get; set; }
        public bool BetStatus { get; set; }
        public int Acumulado { get; set; }
        public byte LastBetNum { get; set; }
        public string LastBetWinner { get; set; }
        public List<(byte, string, int)> Apostas { get; set; }

        public Lottery() { }
        public void Save(Lottery lottery)
        {
            // Cria o diretório se não existir
            if (!Directory.Exists("~/Lottery".MyDir()))
            {
                Directory.CreateDirectory("~/Lottery".MyDir());
            }

            // Adiciona o conversor de tuplas
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new TupleConverter() }
            };

            string json = JsonSerializer.Serialize(lottery, options);
            File.WriteAllText(@"~/Lottery/lotteryData.json".MyDir(), json);
        }
        public Lottery Load()
        {
            string filePath = @"~/Lottery/lotteryData.json".MyDir();
            if (!File.Exists(filePath)) return null;

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new TupleConverter() }
            };

            string json = File.ReadAllText(filePath);

            var ResultJson = JsonSerializer.Deserialize<Lottery>(json, options);

            return ResultJson;
        }
    }
}