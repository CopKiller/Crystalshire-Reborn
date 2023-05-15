using Discord;
using Event_Server.Communication;
using Event_Server.Util;
using System;
using Event_Server.Server;

namespace Event_Server.Network.ClientPacket
{
    public sealed class CpReceiveDiscordMsg : IRecvPacket
    {
        public void Process(byte[] buffer, IConnection connection)
        {
            var msg = new ByteBuffer(buffer);

            // Cria um objeto EmbedBuilder para construir o embed
            EmbedBuilder embedBuilder = new EmbedBuilder();

            // Lê o tipo de mensagem do Discord
            byte dscMsgType = msg.ReadByte();
            if (dscMsgType == 0) { return; }

            // Lê o título da mensagem
            string titleMsg = msg.ReadString();
            if (string.IsNullOrEmpty(titleMsg)) { return; }
            // Define o título do embed
            embedBuilder.Title = titleMsg;

            // Lê o nível
            int lvl = msg.ReadInt32();
            if (lvl == 0) { return; }
            // Lê o status VIP
            byte vip = msg.ReadByte();
            
            if (dscMsgType != (byte)DscMsgType.Chat)
            {
                // Adiciona um campo para exibir o nível
                embedBuilder.AddField("Level", lvl, true);
                // Adiciona um campo para exibir o status VIP
                embedBuilder.AddField("Vip", (vip == 1) ? "Sim" : "Não", true);
            }

            // Tenta ler o texto da mensagem
            var messageText = msg.ReadString();
            if (!string.IsNullOrEmpty(messageText))
            {
                // Define o texto da descrição do embed
                embedBuilder.Description = messageText;
            }

            // Define a URL da miniatura do embed
            embedBuilder.ThumbnailUrl = "https://media4.giphy.com/media/v1.Y2lkPTc5MGI3NjExMmU5ZjQ3Y2U4YzQ5MjFiNDRmZmM4NTc4MDg3YTE5NzYyYmE4MzRhZCZlcD12MV9pbnRlcm5hbF9naWZzX2dpZklkJmN0PXM/wh5frFkevniZXJpYgd/giphy.gif";

            // Envia o embed para o canal do Discord
            Global.DiscordBot.SendEmbed((DscMsgType)dscMsgType, embedBuilder.Build());
        }
    }
}