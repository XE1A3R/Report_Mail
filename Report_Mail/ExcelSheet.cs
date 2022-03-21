#nullable enable
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using TableStyles = OfficeOpenXml.Table.TableStyles;

namespace Report_Mail
{
    public class ExcelSheet
    {
        private readonly ExcelWorksheet _xlWorksheet;
        public ExcelSheet(ExcelPackage excelPackage, string name)
        {
            _xlWorksheet = excelPackage.Workbook.Worksheets.Add(name);
        }
        
        public void InsertData(int locationColumn, int locationRow,
            DataGridView dataGridView1,
            BackgroundWorker backgroundWorker1, List<ExcelSheet>? excelSheet)
        {            
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (excelSheet == null) continue;
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow, i + 1].Value =
                    dataGridView1.Columns[i].HeaderCell.Value.ToString();
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow, i + 1].Style.Font.Size = 12;
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow, i + 1].Style.Font.Bold = true;
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow, i + 1].Style.Border.Top.Style =
                    ExcelBorderStyle.Medium;
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow, i + 1].Style.Border.Bottom.Style =
                    ExcelBorderStyle.Medium;
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow, i + 1].Style.Border.Left.Style =
                    ExcelBorderStyle.Medium;
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow, i + 1].Style.Border.Right.Style =
                    ExcelBorderStyle.Medium;
            }

            if (excelSheet == null) return;
            {
                excelSheet[excelSheet.Count - 1]._xlWorksheet.View.FreezePanes(locationRow + 1, 1);
                var range = excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow, 1,
                    dataGridView1.RowCount + locationRow, dataGridView1.ColumnCount];
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Tables.Add(range,
                    excelSheet[excelSheet.Count - 1]._xlWorksheet.Name
                        .Replace(_xlWorksheet.Name, $"Table{locationRow}")).TableStyle = TableStyles.Medium9;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        backgroundWorker1.ReportProgress(i + j);
                        excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow + 1, j + 1].Value =
                            dataGridView1[j, i].FormattedValue.ToString();
                        excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow + 1, j + 1].Style.Border.Top
                            .Style = ExcelBorderStyle.Thin;
                        excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow + 1, j + 1].Style.Border.Bottom
                            .Style = ExcelBorderStyle.Thin;
                        excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow + 1, j + 1].Style.Border.Left
                            .Style = ExcelBorderStyle.Thin;
                        excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells[locationRow + 1, j + 1].Style.Border.Right
                            .Style = ExcelBorderStyle.Thin;
                    }
                    locationRow++;
                }
                excelSheet[excelSheet.Count - 1]._xlWorksheet.Cells.AutoFitColumns();
            }
        }
    }
}
