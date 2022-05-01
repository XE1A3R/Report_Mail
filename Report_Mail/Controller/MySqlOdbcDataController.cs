using System.Data;
using System.Data.Odbc;
using System.Windows.Forms;
using Report_Mail.Interface;

namespace Report_Mail.Controller
{
    public static class MySqlOdbcDataController 
    {
        public static DataTable Request(string request)
        {
            var s11 = new OdbcConnection(Properties.Settings.Default.S11);
            var adapter = new OdbcDataAdapter(request, s11);
            var dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }
    }
}