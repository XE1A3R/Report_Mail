using System.Collections.Generic;
using OfficeOpenXml;
using Report_Mail.Controller;

namespace Report_Mail.Interface
{
    public interface IExcelSheetController
    {
        ExcelWorksheet Worksheet { get; }
        List<ExcelCellsController> CellsControllers { get; }
    }
}