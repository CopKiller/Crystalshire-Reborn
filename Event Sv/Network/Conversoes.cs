using System;

namespace Data_Server.Network
{
    internal class Conversoes
    {
        public Conversoes() { }
        public bool ByteToBoolean(byte value)
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
    }
}
