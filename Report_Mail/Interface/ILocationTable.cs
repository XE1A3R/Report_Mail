namespace Report_Mail.Interface
{
    public interface ILocationTable
    {
        public string Request { get; set; }
        public int Column { get; set; }
        public int Row { get; set; }
    }
}