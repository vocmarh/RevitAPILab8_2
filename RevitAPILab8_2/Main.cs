using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPILab8_2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Сообщение", "После закрытия данного окна произойдет экспорт файла проекта в формат NWC");

            Document doc = commandData.Application.ActiveUIDocument.Document;
            var nwcOption = new NavisworksExportOptions();
            using (var ts = new Transaction(doc, "export nwc"))
            {
                ts.Start();
                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.nwc", nwcOption);
                ts.Commit();
            }

            TaskDialog.Show("Сообщение", "Экспорт в формат NWC выполнен");
            return Result.Succeeded;
        }
    }
}
