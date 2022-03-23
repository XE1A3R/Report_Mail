using System.Data;

namespace Report_Mail.Interface
{
    public interface IDataController
    {
        DataTable Table { get; }
        void Request(string request);
    }
}