using System.Collections.Generic;

namespace Report_Mail.Interface
{
    public interface IExcelWindowController
    {
        List<IExcelSheetController> SheetControllers { get; }
        void CreateSheet();
        void Save();
    }
}