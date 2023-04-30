using System;

namespace Data_Server.Network
{
    public static class Conversoes
    {
        //public static Conversoes() { }
        public static bool ByteToBoolean(byte value)
        {
            if (value == 0)
            {
                return false;
            }
            else if (value == 1)
            {
                return true;
            }
            else
            {
                throw new ArgumentException("O valor deve ser 0 ou 1.", nameof(value));
            }
        }

        public static byte BooleanToByte(bool value)
        {
            if (!value)
            {
                return 0;
            }
            else if (value)
            {
                return 1;
            }
            else
            {
                throw new ArgumentException("O valor deve ser true ou false.", nameof(value));
            }
        }
    }
}
