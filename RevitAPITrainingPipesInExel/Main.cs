using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITrainingPipesInExel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string pipeInfo = string.Empty;
            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();
            string exelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Трубы.xlsx");

            using (FileStream stream = new FileStream(exelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;
                foreach (Pipe pipe in pipes)
                {

                    //double pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    //double l = Math.Round(UnitUtils.ConvertFromInternalUnits(pipeLength, UnitTypeId.Millimeters), 2);
                    //sheet.SetCellValue(rowIndex, columnIndex: 0, l);

                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString());
                    sheet.SetCellValue(rowIndex, columnIndex: 1, pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble());
                    sheet.SetCellValue(rowIndex, columnIndex: 2, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble());
                    sheet.SetCellValue(rowIndex, columnIndex: 3, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble());
                    rowIndex++;
                }
                workbook.Write(stream);
                workbook.Close();
            }
            System.Diagnostics.Process.Start(exelPath);

            return Result.Succeeded;
        }
    }

}
