using Data_Server.Network;
using Event_Server.Communication;
using Event_Server.Data;
using Event_Server.Network.ServerPacket;
using Event_Server.Util;
using System.Collections.Generic;

namespace Event_Server.Network.ClientPacket
{
    public sealed class CpPing : IRecvPacket  {
        public void Process(byte[] buffer, IConnection connection) {
            // Não recebe o ping de volta, apenas pra não atrapalhar a contagem de cada classe.
        }
    }
}