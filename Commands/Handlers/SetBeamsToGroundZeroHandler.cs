using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Recreates selected structural beams at Level 0 with z-justification set to Top.
    /// The original beams are deleted and replaced with new ones at elevation zero,
    /// preserving the top elevation via Z_OFFSET_VALUE.
    /// </summary>
    public class SetBeamsToGroundZeroHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Find Level 0 (elevation == 0)
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            Element level0 = levels.FirstOrDefault(l => l.Elevation == 0);
            if (level0 == null)
            {
                TaskDialog.Show("Error", "No level with elevation 0 found.");
                return;
            }

            // Filter selected elements to structural framing (beams)
            var beamFilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
            var selectedBeams = new FilteredElementCollector(doc, uidoc.Selection.GetElementIds())
                .WherePasses(beamFilter)
                .ToList();

            if (selectedBeams.Count == 0)
            {
                TaskDialog.Show("Selection", "Please select one or more structural beams.");
                return;
            }

            using (Transaction trans = new Transaction(doc, "Set beams Reference level to Level 0"))
            {
                trans.Start();

                var newBeams = new List<FamilyInstance>();
                var toDelete = new List<ElementId>();

                foreach (Element beamElement in selectedBeams)
                {
                    FamilyInstance bs = beamElement as FamilyInstance;
                    if (bs == null) continue;

                    Curve bsCurve = (bs.Location as LocationCurve)?.Curve;
                    if (bsCurve == null) continue;

                    Line bsLine = bsCurve as Line;
                    if (bsLine == null) continue;

                    var elevationTop = beamElement.get_Parameter(BuiltInParameter.STRUCTURAL_ELEVATION_AT_TOP);
                    if (elevationTop == null) continue;

                    double beamTopHeight = elevationTop.AsDouble();

                    // Create new beam at Level 0
                    FamilyInstance newBeam = doc.Create.NewFamilyInstance(
                        bsLine.Origin, bs.Symbol, level0,
                        StructuralType.Beam);

                    // Set the curve to match the original
                    LocationCurve beamCurve = newBeam.Location as LocationCurve;
                    if (beamCurve != null)
                    {
                        beamCurve.Curve = bsLine;
                    }

                    // Configure the new beam
                    var zJustification1 = newBeam.get_Parameter(BuiltInParameter.Z_JUSTIFICATION);
                    var refLevel1 = newBeam.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                    var zOffsetValue1 = newBeam.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE);
                    var startLevelOffset = newBeam.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION);
                    var endLevelOffset = newBeam.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION);

                    zJustification1?.Set((int)ZJustification.Top);
                    refLevel1?.Set(level0.Id);
                    zOffsetValue1?.Set(beamTopHeight);
                    startLevelOffset?.Set(0);
                    endLevelOffset?.Set(0);

                    newBeams.Add(newBeam);
                    toDelete.Add(beamElement.Id);
                }

                // Delete original beams
                foreach (var id in toDelete)
                {
                    doc.Delete(id);
                }

                trans.Commit();
            }

            TaskDialog.Show("Completed", $"Relocated {selectedBeams.Count} beams to Level 0.");
        }
    }
}
