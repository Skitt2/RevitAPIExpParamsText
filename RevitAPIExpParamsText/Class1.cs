using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPIExpParamsText
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "wallInfo.txt",
                DefaultExt = ".txt"
            };

            string selectedFilePath = string.Empty;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(selectedFilePath))
                return Result.Cancelled;

            List<Wall> allWalls = new FilteredElementCollector(doc)
                 .OfClass(typeof(Wall))
                 .Cast<Wall>()
                 .ToList();

            string allText = string.Empty;
            foreach (var wall in allWalls)
            {
                string wallType = wall.LookupParameter("Марка").AsString();
                string WallVolume = UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble(), UnitTypeId.Meters).ToString();
                allText += $"{wallType}\t{WallVolume}{Environment.NewLine}";
            }

            File.WriteAllText(selectedFilePath, allText);
            return Result.Succeeded;
        }
    }
}
