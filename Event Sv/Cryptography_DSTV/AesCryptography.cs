using System.IO;
using System.Security.Cryptography;

namespace Event_Server.Cryptography
{
    public class AesCryptography
    {
        public CryptographyKeySize KeySize { get; set; }
        public CipherMode CipherMode { get; set; }
        public PaddingMode PaddingMode { get; set; }

        private const int BlockSize = 128;

        public byte[] Encrypt(byte[] bytesToBeEncrypted, byte[] key, byte[] iv)
        {
            byte[] encryptedBytes = null;

            using (var AES = new RijndaelManaged())
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    AES.KeySize = (int)KeySize;
                    AES.BlockSize = BlockSize;
                    AES.Mode = CipherMode;
                    AES.Padding = PaddingMode;

                    AES.Key = key;
                    AES.IV = iv;

                    using (var cs = new CryptoStream(ms, AES.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(bytesToBeEncrypted, 0, bytesToBeEncrypted.Length);
                        cs.FlushFinalBlock();
                        cs.Close();
                    }

                    encryptedBytes = ms.ToArray();
                }
            }

            return encryptedBytes;
        }

        public byte[] Decrypt(byte[] bytesToBeDecrypted, byte[] key, byte[] iv)
        {
            byte[] encryptedBytes = null;
            var success = false;

            try
            {
                using (var AES = new RijndaelManaged())
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        AES.KeySize = (int)KeySize;
                        AES.BlockSize = BlockSize;
                        AES.Mode = CipherMode;
                        AES.Padding = PaddingMode.None;

                        AES.Key = key;
                        AES.IV = iv;

                        using (var cs = new CryptoStream(ms, AES.CreateDecryptor(), CryptoStreamMode.Write))
                        {
                            cs.Write(bytesToBeDecrypted, 0, bytesToBeDecrypted.Length);
                            cs.Close();
                        }

                        success = true;
                        encryptedBytes = ms.ToArray();
                    }
                }
            }
            catch
            {
                success = false;
            }

            if (success)
            {
                return encryptedBytes;
            }

            // Retorna um número fora do range para executar a desconexão.
            return new byte[4] { 255, 255, 255, 255 };
        }

    }
}
