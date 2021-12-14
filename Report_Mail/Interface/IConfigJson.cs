using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

namespace Report_Mail.Interface
{
    public interface IConfigJson
    {
        public List<Xls> Xls { get; set; }
        public List<Mail> Mail { get; set; }
    }
}