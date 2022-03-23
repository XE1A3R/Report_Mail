using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using Report_Mail.Interface;
using Report_Mail.Model;

namespace Report_Mail.Controller
{
    public class ExcelCellsController : ILocationTable
    {
        private readonly Label _label1;
        public bool PrintHeaders { get; set; }
        public bool SmartTable { get; set; }
        public bool FreezePanes { get; set; }
        public uint Size { get; set; }
        public string Request { get; set; }
        public int Column { get; set; }
        public int Row { get; set; }
        private readonly ExcelWorksheet _excelWorksheet;
        private readonly DataTable _table;

        public ExcelCellsController(ILocationTable location, ExcelWorksheet excelWorksheet, Label label1)
        {
            _label1 = label1;
            PrintHeaders = location.PrintHeaders;
            SmartTable = location.SmartTable;
            FreezePanes = location.FreezePanes;
            Size = location.Size;
            Request = location.Request;
            Column = location.Column;
            Row = location.Row;
            _excelWorksheet = excelWorksheet;
            var sql = new MySqlDataController(new DataTable(), label1);
            sql.Request(Request);
            _table = sql.Table;
            InsertData();
        }

        private void InsertData()
        {
            _label1.Text = @"Выгрузка в EXCEL... ";
            if (SmartTable)
                _excelWorksheet.Cells[Row, Column].LoadFromDataTable(_table, PrintHeaders, TableStyles.Medium9);
            else
                _excelWorksheet.Cells[Row, Column].LoadFromDataTable(_table, PrintHeaders);
            _excelWorksheet.Cells[Row, Column].AutoFitColumns();
            if(FreezePanes)
                _excelWorksheet.View.FreezePanes(Row+1,Column);
        }
                        // backgroundWorker.ReportProgress(i + j);
    }
}