using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using Report_Mail.Interface;
using Label = System.Windows.Forms.Label;

namespace Report_Mail.Controller
{
    public class ExcelController : IExcelWindowController
    {
        private readonly IXls _xls;
        private readonly ExcelPackage _excelPackage = new();
        private readonly Label _label1;

        private ExcelWorksheet Worksheet { get; set; }

        public ExcelController(IXls xls, Label label1)
        {
            _xls = xls;
            _label1 = label1;
        }

        public void CreateSheet()
        {
            foreach (var sheet in _xls.Sheets)
            {
                _label1.Invoke((MethodInvoker) delegate
                {
                    _label1.Text = @"Создание нового файла EXCEL...";
                });
                Worksheet = _excelPackage.Workbook.Worksheets.Add(sheet.Name);
                foreach (var location in sheet.Locations)
                {
                    _label1.Invoke((MethodInvoker) delegate
                    {
                        _label1.Text = @"Выполняется запрос... ";
                    });
                    var table = MySqlOdbcDataController.Request(location.Request);
                    _label1.Invoke((MethodInvoker) delegate
                    {
                        _label1.Text = @"Выгрузка в EXCEL... ";
                    });
                    if (location.SmartTable)
                        Worksheet?.Cells[location.Row, location.Column].LoadFromDataTable(table, location.PrintHeaders, TableStyles.Medium9);
                    else
                        Worksheet?.Cells[location.Row, location.Column].LoadFromDataTable(table, location.PrintHeaders);
                    Worksheet?.Cells.AutoFitColumns();
                    if(location.FreezePanes)
                        Worksheet?.View.FreezePanes(location.Row+1,location.Column);
                }
            }
        }

        public void Save()
        {
            _label1.Invoke((MethodInvoker) delegate
            {
                _label1.Text = @"Сохранение EXCEL...";
            });
            var aFile = new FileStream(FileManagerController.GetFile(_xls.Attachments, _xls.Name, _xls.Format), FileMode.Create);
            _excelPackage.SaveAs(aFile);
            _excelPackage.Dispose();
            aFile.Close();
        }
    }
}
