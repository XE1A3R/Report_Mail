using System.Data;
using System.Data.Odbc;
using System.Windows.Forms;
using Report_Mail.Interface;

namespace Report_Mail.Controller
{
    public class MySqlDataController : IDataController
    {
        private readonly Label _label1;
        public DataTable Table { get; }

        public MySqlDataController(DataTable table, Label label1)
        {
            _label1 = label1;
            Table = table;
        }

        public void Request(string request)
        {
            _label1.Text = @"Выполняется процедура...";
            var s11 = new OdbcConnection(Properties.Settings.Default.S11);
            var adapter = new OdbcDataAdapter(request, s11);
            adapter.Fill(Table);
        }
    }
}