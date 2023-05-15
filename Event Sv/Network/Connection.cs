using System;
using System.Collections.Generic;
using System.Net.Sockets;
using Event_Server.Util;
using Event_Server.Communication;
using Event_Server.Network.ServerPacket;
using System.Reflection;
//using Event_Server.Cryptography;
using System.Security.Cryptography;
using System.Net.NetworkInformation;
using System.Net;
using Event_Server.Server;

namespace Event_Server.Network {
    public sealed class Connection : IConnection {
        public static Dictionary<int, Connection> Connections { get; set; } = new Dictionary<int, Connection>();

        // Maior indice da lista de usuários.
        public static int HighIndex { get; set; }

        public string UniqueKey { get; set; }
        public bool Connected { get; set; }

        public string IpAddress { get; set; }

        private TcpClient Client;
        private ByteBuffer msg;
        private static int pingTick;

        private SpPing sPing = new SpPing();

        //private byte[] decryptedBytes;
        //readonly ByteBuffer decrypted_msg;
        //readonly AesCryptography aes;

        public Connection(TcpClient tcpClient, string ipAddress, string uniqueKey) {

            msg = new ByteBuffer();

            IpAddress = ipAddress;
            UniqueKey = uniqueKey;

            Client = tcpClient;

            Connected = true;

            // Verifica se tem algum client conectado, caso esteja sai do método,
            // limita apenas 1 conexão no eventserver, caso queira alterar pra receber
            // mais de uma conexão comente a condição abaixo!!!
            if (Connections.Count > 0)
            {
                Disconnect(); // Adiciona o comando Close() na conexão rejeitada
                Global.WriteLog(LogType.System, $"{IpAddress} {uniqueKey} Foi rejeitado pois já tem uma conexão ativa!", LogColor.Red);
                return;
            }

            Add(this);

            //Global.WriteLog(LogType.System, $"{ipAddress} Key {uniqueKey} is connected", LogColor.Coral);
            Global.WriteLog(LogType.System, $"{ipAddress} Key {uniqueKey} Main Server is connected", LogColor.Green);
        }

        public Connection()
        {
        }

        public void SendPing()
        {
            if (Environment.TickCount >= Connection.pingTick)
            {
                if (Connection.HighIndex > 0)
                {

                    SpPing.SendPacket();
                }
                Connection.pingTick = Environment.TickCount + Constants.PingTime;
            }
        }

        public void Disconnect() {
            Client.Close();
            Connections.Remove(Connection.HighIndex);
            Connection.HighIndex = 0;
            Connected = false;
        }

        public void ReceiveData() {
            if (Client.Client == null) {
                return;
            }

            if (Client.Available > 0) {
                var size = Client.Available;
                byte[] buffer = new byte[size];

                if (Client.Client.Poll(Constants.ReceiveTimeOut, SelectMode.SelectRead)) {
                    try {
                        // Recebe os primeiros dados.
                        Client.Client.Receive(buffer, size, SocketFlags.None);

                        // Escreve o buffer.
                        msg.Write(buffer);
                        var pSize = msg.ReadInt32(false);

                        // Enquanto a mensagem nao chegar por completo, lê os dados e adiciona no buffer.
                        while (msg.Count() - 4 < pSize) {
                            if (Client.Available > 0) {
                                buffer = new byte[Client.Available];

                                Client.Client.Receive(buffer, Client.Available, SocketFlags.None);

                                msg.Write(buffer);
                            }
                        }
                    }
                    catch (SocketException ex) {
                        Global.WriteLog(LogType.System, $"Receive Data Error: Class {GetType().Name}", LogColor.Red);
                        Global.WriteLog(LogType.System, $"Message: {ex.Message}", LogColor.Red);
                        Disconnect();
                        Global.WriteLog(LogType.System, $"Main Server is desconnected", LogColor.Blue);
                        return;
                    }
                }
                else {
                    Disconnect();
                    Global.WriteLog(LogType.System, $"Main Server is desconnected", LogColor.Blue);
                    return;
                }

                int pLength = 0;

                Global.WriteLog(LogType.Debug, $"Receive -> {IpAddress} Buffer Size: {msg.Length()}", LogColor.Black);

                if (msg.Length() >= 4) {
                    pLength = msg.ReadInt32(false);

                    if (pLength < 0) {
                        return;
                    }
                }

                //int encrypted_length;
                //byte key;
                //int keyIndex;
                //byte iv;
                //int ivIndex;

                while (pLength > 0 && pLength <= msg.Length() - 4) {
                    if (pLength <= msg.Length() - 4) {
                        // Remove the first packet (Size of Packet).
                        msg.ReadInt32();

                        //  decrypt
                        //encrypted_length = msg.ReadInt32();
                        //key = msg.ReadByte();
                        //keyIndex = msg.ReadByte();
                        //iv = msg.ReadByte();
                        //ivIndex = msg.ReadByte();
                        //var _key = ConnectionPassword.CreateKey(CryptographyKeyType.Key, keyIndex, key);
                        //var _iv = ConnectionPassword.CreateKey(CryptographyKeyType.Iv, ivIndex, iv);
                        //decryptedBytes = aes.Decrypt(msg.ReadBytes(encrypted_length), _key, _iv);
                        //.Write(decryptedBytes);

                        // Remove the header.
                        var header = msg.ReadInt32();
                        // Decrease 4 bytes of header.
                        pLength -= 4;

                        if (OpCode.RecvPacket.ContainsKey(header)) {
                            ((IRecvPacket)Activator.CreateInstance(OpCode.RecvPacket[header])).Process(msg.ReadBytes(pLength), this);
                        }
                        else {
                            Global.WriteLog(LogType.System, $"Header: {header} was not found", LogColor.Red);
                        }
                    }

                    pLength = 0;

                    if (msg.Length() >= 4) {
                        pLength = msg.ReadInt32(false);

                        if (pLength < 0) {
                            return;
                        }
                    }
                }

                msg.Clear();
            }
        }

        public void Send(ByteBuffer msg, string className) {
            var buffer = new byte[msg.Length() + 4];
            var values = BitConverter.GetBytes(msg.Length());

            Array.Copy(values, 0, buffer, 0, 4);
            Array.Copy(msg.ToArray(), 0, buffer, 4, msg.Length());

            if (Client.Client.Poll(Constants.SendTimeOut, SelectMode.SelectWrite)) {
                try {
                    Client.Client.Send(buffer, SocketFlags.None);
                }
                catch (SocketException ex) {
                    if (ex.ErrorCode > 0) {

                    }

                     Global.WriteLog(LogType.System, $"Send Data Error: Class {className}", LogColor.Red);
                     Global.WriteLog(LogType.System, $"Message: {ex.Message}", LogColor.Red);
                    Disconnect();
                }
            }
            else {
                Disconnect();
            }
        }

        public static void Remove(int index) {
            var connection = Connections[index];

            if (Connections.ContainsKey(index)) {
                Connections.Remove(index);
                connection.Connected = false;
            }
        }

        private static void Add(Connection connection) {

            if (Connections.Count == 0) {
                ++HighIndex;
                Connections.Add(HighIndex, connection);
            }
        }     
    }
} 