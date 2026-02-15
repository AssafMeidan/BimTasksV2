using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Copies all walls from a user-selected linked Revit model into the current document.
    /// Uses the link's transform to place walls at correct coordinates.
    /// Selects the newly copied walls in the UI after completion.
    /// </summary>
    public class CopyLinkedWallsHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Prompt user to select a Revit Link Instance
                Reference linkRef;
                try
                {
                    linkRef = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        new LinkSelectionFilter(),
                        "Select the Revit Link instance to copy walls from");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return; // User pressed Esc
                }

                if (linkRef == null) return;

                RevitLinkInstance linkInstance = doc.GetElement(linkRef) as RevitLinkInstance;
                Document linkDoc = linkInstance?.GetLinkDocument();

                if (linkDoc == null)
                {
                    TaskDialog.Show("Error", "The linked document is not loaded in the project.");
                    return;
                }

                // Collect all walls from the linked document
                var wallsInLink = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .ToElementIds()
                    .ToList();

                if (wallsInLink.Count == 0)
                {
                    TaskDialog.Show("Info", "No walls found in the selected link.");
                    return;
                }

                // Get the transform to place walls at correct coordinates
                Transform linkTransform = linkInstance.GetTotalTransform();

                // Copy the walls
                var newWallIds = new List<ElementId>();

                using (Transaction t = new Transaction(doc, "Copy Linked Walls"))
                {
                    t.Start();

                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                        linkDoc,
                        wallsInLink,
                        doc,
                        linkTransform,
                        new CopyPasteOptions());

                    newWallIds.AddRange(copiedIds);
                    t.Commit();
                }

                // Select the new walls in the UI
                if (newWallIds.Count > 0)
                {
                    uidoc.Selection.SetElementIds(newWallIds);
                    TaskDialog.Show("Success", $"Successfully copied and selected {newWallIds.Count} walls.");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "CopyLinkedWallsHandler failed");
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
            }
        }
    }
}
