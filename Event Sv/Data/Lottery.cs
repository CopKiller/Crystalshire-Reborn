namespace Event_Server.Data {
    public sealed class Lottery {
        public string AccountName { get; set; }
        public string Password { get; set; }
        public int AccountId { get; set; }
        public string Email { get; set; }
        public byte Banned { get; set; }
        public int UserGroup { get; set; }
        public byte ServiceId { get; set; }
    }
}