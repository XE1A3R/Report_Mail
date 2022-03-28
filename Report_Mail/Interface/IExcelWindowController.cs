using System.Collections.Generic;
using Report_Mail.Controller;

namespace Report_Mail.Interface
{
    public interface IExcelWindowController
    {
        void CreateSheet();
        void Save();
    }
}