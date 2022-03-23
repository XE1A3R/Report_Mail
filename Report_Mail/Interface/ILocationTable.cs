namespace Report_Mail.Interface
{
    public interface ILocationTable
    {
        public bool PrintHeaders { get; set; }
        public bool SmartTable { get; set; }
        public bool FreezePanes { get; set; }
        public uint Size { get; set; }
        public string Request { get; set; }
        public int Column { get; set; }
        public int Row { get; set; }
    }
}