#nullable enable
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Report_Mail.Interface;
using Label = System.Windows.Forms.Label;

namespace Report_Mail.Controller
{
    public class ExcelWindowController : IExcelWindowController
    {
        private readonly IXls _xls;
        private readonly ExcelPackage _excelPackage = new();
        private readonly Label _label1;
        private List<ExcelSheetController> SheetControllers { get; } = new();

        public ExcelWindowController(IXls xls, Label label1)
        {
            _xls = xls;
            _label1 = label1;
        }

        public void CreateSheet()
        {
            foreach (var sheet in _xls.Sheets)
            {
                _label1.Text = @"Создание нового файла EXCEL...";
                SheetControllers.Add(new ExcelSheetController(sheet, _excelPackage, _label1));
            }
        }

        public void Save()
        {
            _label1.Text = @"Сохранение EXCEL...";
            var aFile = new FileStream(FileManagerController.GetFile(_xls.Attachments, _xls.Name, _xls.Format), FileMode.Create);
            _excelPackage.SaveAs(aFile);
            _excelPackage.Dispose();
            aFile.Close();
        }
    }
}
