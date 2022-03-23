using System.Data;
using OfficeOpenXml;

namespace Report_Mail.Interface
{
    public interface ITableStyle
    {
        void Add(ExcelWorksheet worksheet, ILocationTable location, DataTable table);
    }
}