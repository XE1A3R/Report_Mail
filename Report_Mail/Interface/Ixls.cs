using System.Collections.Generic;

namespace Report_Mail.Interface
{
    public interface IXls
    {
        public string name { get; set; }
        public List<Sheet> Sheets { get; set; }
        public string Attachments { get; set; }
        public string Format { get; set; }
    }
}
