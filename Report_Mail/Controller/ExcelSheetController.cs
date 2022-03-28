#nullable enable
using System.Collections.Generic;
using System.Windows.Forms;
using OfficeOpenXml;
using Report_Mail.Interface;

namespace Report_Mail.Controller
{
    public class ExcelSheetController
    {
        private readonly ISheet _sheet;
        private readonly Label _label1;
        private ExcelWorksheet Worksheet { get; }

        private List<ExcelCellsController> CellsControllers { get; } = new();

        public ExcelSheetController(ISheet sheet, ExcelPackage excelPackage, Label label1)
        {
            _sheet = sheet;
            _label1 = label1;
            Worksheet = excelPackage.Workbook.Worksheets.Add(sheet.Name);
            InsertCell();
        }

        private void InsertCell()
        {
            _label1.Text = @"Выгрузка в EXCEL... ";
            if (_sheet.Locations == null) return;
            foreach (var location in _sheet.Locations)
            {
                CellsControllers.Add(new ExcelCellsController(location, Worksheet, _label1));
            }
        }
    }
}