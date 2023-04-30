using System;
using System.Security.Cryptography;

namespace Event_Server.Cryptography
{
    public static class Hash
    {
        /// <summary>
        /// Retorna um hash a partir dos dados fornecidos.
        /// </summary>
        /// <param name="data">dados a serem computados</param>
        /// <param name="length">tamanho do retorno</param>
        /// <returns></returns>
        public static byte[] Compute(byte[] data, CryptographyKeyLength length, CryptographyKeyType keyType)
        {
            var hash = new byte[(int)length];
            var copy = new byte[data.Length];

            Array.Copy(data, 0, copy, 0, data.Length);

            if (keyType == CryptographyKeyType.Key)
            {
                Array.Reverse(copy);
            }

            var sha = new SHA256Managed();
            var buffer = sha.ComputeHash(copy);

            sha.Dispose();

            Array.Copy(buffer, 0, hash, 0, (int)length);

            return hash;
        }
    }
}
