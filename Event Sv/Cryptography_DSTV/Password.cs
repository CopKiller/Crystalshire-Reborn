
namespace Event_Server.Cryptography
{
    public class Password
    {
        public string HexPassword { get; set; }
        public byte[] Key { get; set; }
        public byte[] Iv { get; set; }
    }
}
