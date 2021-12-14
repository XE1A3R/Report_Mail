using System.Collections.Generic;

namespace Report_Mail.Interface
{
    public interface ISheet
    {
        public string Name { get; set; }
        public List<LocationTable> Locations {get; set; }
    }
}