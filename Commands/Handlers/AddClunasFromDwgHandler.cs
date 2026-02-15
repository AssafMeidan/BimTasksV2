using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers;
using BimTasksV2.Models.Dwg;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Places structural foundation (pile) family instances at DWG arc centers.
    /// The user selects an imported DWG, arcs on the "fnd" layer are parsed,
    /// and a structural foundation family named "Auto" is placed at each arc center
    /// with the diameter set to the arc's diameter.
    /// </summary>
    public class AddClunasFromDwgHandler : ICommandHandler
    {
        private const string FamilyTypeName = "Auto";

        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Prompt user to select the DWG import
                Reference r = uidoc.Selection.PickObject(ObjectType.Element,
                    "Select a DWG import instance containing pile arcs");

                ImportInstance myDWG = doc.GetElement(r) as ImportInstance;
                if (myDWG == null)
                {
                    TaskDialog.Show("Error", "Selected element is not a DWG import instance.");
                    return;
                }

                // Parse arcs from the DWG
                var dwgWrapper = new DwgWrapper(myDWG);
                if (dwgWrapper.GeometryArcs.Count == 0)
                {
                    TaskDialog.Show("Info", "No arcs found on the 'fnd' layer in the selected DWG.");
                    return;
                }

                // Find the foundation family type
                var symbols = RevitFilterHelper.GetElementsOfType(doc, typeof(ElementType),
                    BuiltInCategory.OST_StructuralFoundation);

                ElementType foundationType = symbols
                    .WhereElementIsElementType()
                    .FirstOrDefault(x => x.Name == FamilyTypeName) as ElementType;

                if (foundationType == null)
                {
                    TaskDialog.Show("Error",
                        $"Structural foundation family type '{FamilyTypeName}' not found in the project.");
                    return;
                }

                FamilySymbol symbol = foundationType as FamilySymbol;
                if (symbol == null)
                {
                    TaskDialog.Show("Error",
                        $"'{FamilyTypeName}' is not a valid FamilySymbol.");
                    return;
                }

                // Get the active view's level
                Level level = uidoc.ActiveView.GenLevel;
                if (level == null)
                {
                    TaskDialog.Show("Error", "Active view does not have an associated level.");
                    return;
                }

                int placedCount = 0;

                using (Transaction t = new Transaction(doc, "Place Piles from DWG"))
                {
                    t.Start();

                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                    }

                    foreach (var arc in dwgWrapper.GeometryArcs)
                    {
                        FamilyInstance fi = doc.Create.NewFamilyInstance(
                            arc.Center, symbol, level, StructuralType.Footing);

                        // Set diameter parameter to arc diameter
                        Parameter prDiameter = fi.LookupParameter("Diameter");
                        if (prDiameter != null && !prDiameter.IsReadOnly)
                        {
                            prDiameter.Set(arc.Radius * 2);
                        }

                        placedCount++;
                    }

                    t.Commit();
                }

                TaskDialog.Show("Completed", $"Placed {placedCount} structural foundations from DWG arcs.");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User pressed Esc
            }
            catch (Exception ex)
            {
                Log.Error(ex, "AddClunasFromDwgHandler failed");
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
            }
        }
    }
}
