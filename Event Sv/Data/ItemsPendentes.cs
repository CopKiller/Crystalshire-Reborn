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
    public sealed class ItemsPendentesData
    {
        public List<(string, int, int, string)> Items { get; set; }

        public ItemsPendentesData()
        {
            Items = new List<(string, int, int, string)>();
        }
    }
    public sealed class ItemsPendentes
    {
        //Directory
        private const string DirItemsPendentes = @"~\ItemsPendentes";

        public ItemsPendentes() { }
        public void Save(ItemsPendentesData items)
        {
            // Cria o diretório se não existir
            if (!Directory.Exists(DirItemsPendentes.MyDir()))
            {
                Directory.CreateDirectory(DirItemsPendentes.MyDir());
            }

            // Adiciona o conversor de tuplas
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new TupleConverterSII() }
            };

            string json = JsonSerializer.Serialize(items, options);
            File.WriteAllText($"{DirItemsPendentes.MyDir()}" + "/ItemsPendentes.json", json);
        }
        public ItemsPendentesData Load()
        {
            string filePath = $"{DirItemsPendentes.MyDir()}/ItemsPendentes.json";

            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Converters = { new TupleConverterSII() }
                };

                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    var ResultJson = JsonSerializer.Deserialize<ItemsPendentesData>(json, options);
                    return ResultJson;
                }
                else
                {
                    return new ItemsPendentesData();
                }
            }
            catch (Exception ex)
            {
                Global.WriteLog(LogType.System, "Ocorreu um erro: " + ex.Message, LogColor.Red);
                return new ItemsPendentesData();
            }
        }
    }
}