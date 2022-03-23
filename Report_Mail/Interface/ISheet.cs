using System.Collections.Generic;
using Report_Mail.Model;

namespace Report_Mail.Interface
{
    public interface ISheet
    {
        public string Name { get; set; }
        public List<LocationTable> Locations {get; set; }
    }
}