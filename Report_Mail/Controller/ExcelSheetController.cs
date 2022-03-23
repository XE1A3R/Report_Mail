#nullable enable
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using Report_Mail.Interface;
using Report_Mail.Model;

namespace Report_Mail.Controller
{
    public class ExcelSheetController : IExcelSheetController, ISheet
    {
        private readonly Label _label1;
        public string Name { get; set; }
        public List<LocationTable>? Locations { get; set; } = new();
        public ExcelWorksheet Worksheet { get; }
        public List<ExcelCellsController> CellsControllers { get; } = new();

        public ExcelSheetController(Sheet sheet, ExcelPackage excelPackage, Label label1)
        {
            _label1 = label1;
            Name = sheet.Name;
            Locations?.AddRange(sheet.Locations ?? throw new InvalidOperationException());
            Worksheet = excelPackage.Workbook.Worksheets.Add(sheet.Name);
            InsertCell();
        }

        private void InsertCell()
        {
            _label1.Text = @"Выгрузка в EXCEL... ";
            if (Locations == null) return;
            foreach (var location in Locations)
            {
                CellsControllers.Add(new ExcelCellsController(location, Worksheet, _label1));
            }
        }
    }
}