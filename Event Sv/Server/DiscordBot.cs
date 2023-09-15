using Discord.WebSocket;
using Discord;
using System;
using System.Threading.Tasks;
using Event_Server.Communication;
using Event_Server.Util;
using System.Collections.Generic;

namespace Event_Server.Server
{
    public enum DscMsgType
    {
        Entrou = 1,
        Levelup,
        Chat,
        Death,

        Count
    }
    public class DiscordChannels
    {
        private static Dictionary<DscMsgType, ulong> channelIds = new Dictionary<DscMsgType, ulong>()
    {
        { DscMsgType.Entrou, 1107050001017872415 },
        { DscMsgType.Levelup, 1107368932479881296 },
        { DscMsgType.Chat, 1107370214028492932 },
        { DscMsgType.Death, 1107370423513010287 }
    };

        public static ulong GetChannelId(DscMsgType messageType)
        {
            return channelIds[messageType];
        }
    }

    public class DiscordBot
    {
        // Token do bot do Discord
        private const string botToken = "MTEwNzA1MDg0MTc5OTY1OTUzMQ.GUpVYH.oA_EYkryEbj6DEfE8ncQwVhIhZ6pGu2io2hHIo";

        private DiscordSocketClient _client;
        private string _token = botToken;

        public DiscordBot()
        {
            _client = new DiscordSocketClient();
        }

        public async Task Start()
        {

            await _client.LoginAsync(TokenType.Bot, _token);
            await _client.StartAsync();
        }

        public async void SendEmbed(DscMsgType channelID, Embed receiveEmbed)
        {
            var channel = _client.GetChannel(DiscordChannels.GetChannelId(channelID)) as SocketTextChannel;

            if (channel != null)
            {
                await channel.SendMessageAsync(embed: receiveEmbed);
            }
        }

    }





    // Receber mensagens do discord aqui no C#

    //class DiscordBot
    //{
    //    private DiscordSocketClient _client;

    //    public async Task RunAsync()
    //    {
    //        _client = new DiscordSocketClient();

    //        _client.Log += Log;

    //        _client.MessageReceived += MessageReceived;

    //        await _client.LoginAsync(TokenType.Bot, "SEU_TOKEN");

    //        await _client.StartAsync();

    //        await Task.Delay(-1);
    //    }

    //    private Task Log(LogMessage log)
    //    {
    //        Console.WriteLine(log.Message);
    //        return Task.CompletedTask;
    //    }

    //    private async Task MessageReceived(SocketMessage message)
    //    {
    //        // Verifique se a mensagem foi enviada por um usuário e não é uma mensagem do bot
    //        if (message.Author.IsBot)
    //            return;

    //        // Faça o processamento da mensagem recebida
    //        // Exemplo: exiba o conteúdo da mensagem no console
    //        Console.WriteLine($"Nova mensagem recebida: {message.Content}");
    //    }
    //}
}
