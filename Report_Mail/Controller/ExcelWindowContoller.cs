#nullable enable
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using Report_Mail.Interface;
using Report_Mail.Model;

namespace Report_Mail.Controller
{
    public class ExcelWindowController : IExcelWindowController, IXls
    {
        public string Name { get; set; }
        public List<Sheet> Sheets { get; set; } = new();
        public string Attachments { get; set; }
        public string Format { get; set; }
        private readonly ExcelPackage _excelPackage = new();
        private readonly Label _label1;
        public List<IExcelSheetController> SheetControllers { get; } = new();

        public ExcelWindowController(IXls xls, Label label1)
        {
            _label1 = label1;
            Name = xls.Name;
            Sheets?.AddRange(xls.Sheets);
            Attachments = xls.Attachments;
            Format = xls.Format;
        }

        public void CreateSheet()
        {
            foreach (var sheet in Sheets)
            {
                _label1.Text = @"Создание нового файла EXCEL...";
                SheetControllers.Add(new ExcelSheetController(sheet, _excelPackage, _label1));
            }
        }

        public void Save()
        {
            _label1.Text = @"Сохранение EXCEL...";
            var aFile = new FileStream(FileManagerController.GetFile(Attachments, Name, Format), FileMode.Create);
            _excelPackage.SaveAs(aFile);
            _excelPackage.Dispose();
            aFile.Close();
        }

    }
}