using Event_Server.Data;
using Event_Server.Communication;
using Event_Server.Util;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Event_Server.Network.ServerPacket;

namespace Event_Server.Network.ClientPacket
{
    public sealed class CpReceiveItemsPendentes : IRecvPacket
    {
        enum Operacao
        {
            None = 0,
            Save,
            Delete,
            Request
        }
        public void Process(byte[] buffer, IConnection connection)
        {
            var msg = new ByteBuffer(buffer);

            var ItemsPendentes = new ItemsPendentes();

            var ItemsLoad = ItemsPendentes.Load();

            var Operacao = msg.ReadByte();

            switch (Operacao)
            {
                case 1: //Operação pra Salvar um novo item pendente!
                    {
                        //Cria uma Tupla de items a serem recebidos pra depois passar pro método!
                        var items = (msg.ReadString(), msg.ReadInt32(), msg.ReadInt32(), msg.ReadString());

                        ItemsLoad.Items.Add(items);

                        ItemsPendentes.Save(ItemsLoad);

                        break;

                    }
                case 2: //Operação pra Excluir um item que estava pendente e ja foi entregue!
                    {
                        var Nome = msg.ReadString().Trim();
                        var ItemID = msg.ReadInt32();
                        var ItemValue = msg.ReadInt32();

                        //Consegue buscar uma area da lista que contenham os 3 valores recebidos
                        //Em seguida ele vai pegar este indice e remover!
                        int Indice = ItemsLoad.Items.FindIndex(t => t.Item1 == Nome &&
                                                       t.Item2 == ItemID &&
                                                       t.Item3 == ItemValue);
                        if (Indice != -1)
                        {
                            ItemsLoad.Items.RemoveAt(Indice);

                            ItemsPendentes.Save(ItemsLoad);
                        }
                        break;
                    }
                case 3: //Operação pra Solicitar um item pendente e quantidade apartir do nome do jogador!
                    {
                        var Nome = msg.ReadString().Trim();

                        // Cria uma nova lista contendo apenas as tuplas que contêm a string "item1"
                        List<(string, int, int, string)> novaLista = ItemsLoad.Items.Where(t => t.Item1 == Nome).ToList();
                        SpItemsPendentes.SendPacket(ItemsLoad);

                        break;
                    }
                default: //Caso receba uma operação nula ou com valores indevidos!
                    {
                        Global.WriteLog(LogType.System, $"Operação recebida invalida: {Operacao}", LogColor.Red);
                        break;
                    }
            }
        }
    }
}