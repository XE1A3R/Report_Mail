using System.Collections.Generic;
using Report_Mail.Model;

namespace Report_Mail.Interface
{
    public interface IXls
    {
        public string Name { get; set; }
        public List<Sheet> Sheets { get; set; }
        public string Attachments { get; set; }
        public string Format { get; set; }
    }
}
