using System.Data;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using Report_Mail.Interface;

namespace Report_Mail.Controller
{
    public class ExcelCellsController
    {
        private readonly ILocationTable _location;
        private readonly Label _label1;
        private readonly ExcelWorksheet _excelWorksheet;
        private readonly DataTable _table;

        public ExcelCellsController(ILocationTable location, ExcelWorksheet excelWorksheet, Label label1)
        {
            _location = location;
            _label1 = label1;
            _excelWorksheet = excelWorksheet;
            var sql = new MySqlDataController(new DataTable(), label1);
            sql.Request(_location.Request);
            _table = sql.Table;
            InsertData();
        }

        private void InsertData()
        {
            _label1.Text = @"Выгрузка в EXCEL... ";
            if (_location.SmartTable)
                _excelWorksheet.Cells[_location.Row, _location.Column].LoadFromDataTable(_table, _location.PrintHeaders, TableStyles.Medium9);
            else
                _excelWorksheet.Cells[_location.Row, _location.Column].LoadFromDataTable(_table, _location.PrintHeaders);
            _excelWorksheet.Cells[_location.Row, _location.Column].AutoFitColumns();
            if(_location.FreezePanes)
                _excelWorksheet.View.FreezePanes(_location.Row+1,_location.Column);
        }
                        // backgroundWorker.ReportProgress(i + j);
    }
}