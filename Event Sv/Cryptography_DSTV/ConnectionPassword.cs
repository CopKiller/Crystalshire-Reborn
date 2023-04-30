using System;

namespace Event_Server.Cryptography
{
    public static class ConnectionPassword
    {
        public static byte[] Key { get; set; }
        public static byte[] Iv { get; set; }

        public static void InitializePassword()
        {
            Key = new byte[(int)CryptographyKeyLength.Key];
            Iv = new byte[(int)CryptographyKeyLength.Iv];

            Random random = new Random(75460013); // passa a semente "75460013" para o gerador de números aleatórios

            for (var i = 0; i < (int)CryptographyKeyLength.Key; i++)
            {
                Key[i] = (byte)(random.Next(0, 257));
            }

            for (var i = 0; i < (int)CryptographyKeyLength.Iv; i++)
            {
                Iv[i] = (byte)(random.Next(0, 257));
            }
        }

        public static byte[] CreateKey(CryptographyKeyType keyType, int index, byte value)
        {
            byte[] array = null;

            if (keyType == CryptographyKeyType.Key)
            {
                array = Key;
            }
            else if (keyType == CryptographyKeyType.Iv)
            {
                array = Iv;
            }

            return CreateArray(array, index, value);
        }

        private static byte[] CreateArray(byte[] array, int index, byte value)
        {
            var dest = new byte[array.Length];

            Array.Copy(array, 0, dest, 0, array.Length);

            dest[index] = value;

            return dest;
        }
    }
}