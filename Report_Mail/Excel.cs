#nullable enable
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using Report_Mail.Interface;

namespace Report_Mail
{
    public class Excel
    {
        private readonly ExcelPackage _package = new();

        public List<ExcelSheet>? ExcelSheet { get; } = new();
        
        public void CreateSheet(string name)
        {
            ExcelSheet?.Add(new ExcelSheet(_package, name));
        }

        public void Save(string file)
        {
            if (FileExists(file))
                FileDelete(file);
            Directory.CreateDirectory(Environment.CurrentDirectory);
            var aFile = new FileStream(file, FileMode.Create);
            _package.SaveAs(aFile);
            _package.Dispose();
            aFile.Close();
        }

        public void Save(Xls xls)
        {
            if (!Directory.Exists(xls.Attachments))
                Directory.CreateDirectory(xls.Attachments); 
            else if (FileExists(xls)) 
                FileDelete(xls);
            var aFile = new FileStream(GetFile(xls), FileMode.Create);
                _package.SaveAs(aFile);
                _package.Dispose();
                aFile.Close();
        }

        private void FileDelete(string file)
        {
            File.Delete(file);
        }

        private void FileDelete(IXls xls)
        {
            File.Delete(GetFile(xls));
        }

        private static bool FileExists(string file)
        {
            return File.Exists(file);
        }
        
        private static bool FileExists(IXls xls)
        {
            return File.Exists(GetFile(xls));
        }

        public static void CreateTable(string request, ref DataGridView dataGridView1)
        {
            var cnS11 = new OdbcConnection(Properties.Settings.Default.S11);
            var adapter1 = new OdbcDataAdapter(request, cnS11);
            var table1 = new DataTable();
            adapter1.Fill(table1);
            dataGridView1.DataSource = table1;
        }

        private static string GetFile(IXls xls)
        {
            return @$"{xls.Attachments}\{xls.name}.{xls.Format}";
        }
    }
}