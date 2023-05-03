using System.Collections.Generic;
using System.Text.Json;
using System.IO;
using Event_Server.Util;
using Data_Server.Util;
using Event_Server.Network;
using Event_Server.Communication;
using System;

namespace Event_Server.Data
{
    public sealed class ItemsPendentes
    {
        public List<(string, int, int)> Items { get; set; }

        public ItemsPendentes() { }
        public void Save(ItemsPendentes items)
        {
            // Cria o diretório se não existir
            if (!Directory.Exists("~/ItemsPendentes".MyDir()))
            {
                Directory.CreateDirectory("~/ItemsPendentes".MyDir());
            }

            // Adiciona o conversor de tuplas
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new TupleConverterSII() }
            };

            string json = JsonSerializer.Serialize(items, options);
            File.WriteAllText(@"~/ItemsPendentes/ItemsPendentes.json".MyDir(), json);
        }
        public ItemsPendentes Load()
        {
            string filePath = @"~/ItemsPendentes/ItemsPendentes.json".MyDir();

            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Converters = { new TupleConverterSII() }
                };

                string json = File.ReadAllText(filePath);

                var ResultJson = JsonSerializer.Deserialize<ItemsPendentes>(json, options);
                return ResultJson;
            }
            catch(Exception ex)
            {
                Global.WriteLog(LogType.System, "Ocorreu um erro: " + ex.Message, LogColor.Red);
                return null;
            }
        }
    }
}