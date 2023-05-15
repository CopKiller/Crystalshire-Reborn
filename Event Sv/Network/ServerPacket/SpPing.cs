using System;
using System.Collections.Generic;
using Event_Server.Communication;
using System.Net.NetworkInformation;
using Event_Server.Data;

namespace Event_Server.Network.ServerPacket {
    public sealed class SpPing : SendPacket {

        public SpPing() {
            msg = new ByteBuffer();
            msg.Write(OpCode.SendPacket[GetType()]);
        }

        // Exemplo de como enviar uma packet pro client sem precisar do servidor principal acioná-la!
        public static void SendPacket()
        {
            if (Connection.HighIndex > 0)
            {
                new SpPing().Send(Connection.Connections[Connection.HighIndex]);
            }
        }
    }
}