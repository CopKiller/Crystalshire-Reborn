namespace Event_Server.Communication
{
    public static class Constants
    {
        /// <summary>
        /// Tempo limite de leitura em microsegundos.
        /// </summary>
        public const int ReceiveTimeOut = 15 * 1000 * 1000;

        /// <summary>
        /// Tempo limite de envio em microsegundos.
        /// </summary>
        public const int SendTimeOut = 15 * 1000 * 1000;

        /// <summary>
        /// Tempo entre cada ping.
        /// </summary>
        public const int PingTime = 15000;

        /// <summary>
        /// Dados de Conexão.
        /// </summary>
        public const string IP = "127.0.0.1";
        public const int PORT = 7004;

        // Events
        public const byte MaxBets = 100;

        // Directories
        public const string DirLottery = @"~\Lottery";
        public const string DirItemsPendentes = @"~\ItemsPendentes";
    }
}