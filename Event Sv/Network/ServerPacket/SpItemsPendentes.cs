using System.Collections.Generic;
using Data_Server.Network;
using Event_Server.Communication;
using Event_Server.Data;

namespace Event_Server.Network.ServerPacket
{
    public sealed class SpItemsPendentes : SendPacket
    {
        public SpItemsPendentes(ItemsPendentesData Items)
        {

            msg = new ByteBuffer();
            msg.Write(OpCode.SendPacket[GetType()]);
            

            if (Items.Items.Count > 0)
            {
                //Envia a quantidade de items do array
                msg.Write(Items.Items.Count);

                for (int i = 0; i < Items.Items.Count; i++)
                {
                    //Nome do jogador
                    msg.Write(Items.Items[i].Item1.Trim());
                    //Item ID
                    msg.Write(Items.Items[i].Item2);
                    //Quantidade
                    msg.Write(Items.Items[i].Item3);
                    //Mensagem
                    msg.Write(Items.Items[i].Item4);
                }
            }
        }

        // Exemplo de como enviar uma packet pro client sem precisar do servidor principal acioná-la!
        public static void SendPacket(ItemsPendentesData Items)
        {
            if (Connection.HighIndex > 0)
            {
                new SpItemsPendentes(Items).Send(Connection.Connections[Connection.HighIndex]);
            }
        }
    }
}