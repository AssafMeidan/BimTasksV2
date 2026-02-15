using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using BimTasksV2.Commands.Infrastructure;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Creates a new instance of the selected family type at a picked point.
    /// Picks an existing element, determines its type, then requests placement of that type.
    /// </summary>
    public class CreateSelectedItemHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // Pick an element to copy its type
                Reference r = uidoc.Selection.PickObject(ObjectType.Element, "Select an element to create another instance of its type.");
                Element pickedElement = doc.GetElement(r);
                if (pickedElement == null)
                {
                    TaskDialog.Show("BimTasksV2", "No element selected.");
                    return;
                }

                var typeId = pickedElement.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId)
                {
                    TaskDialog.Show("BimTasksV2", "Selected element has no type.");
                    return;
                }

                ElementType elementType = doc.GetElement(typeId) as ElementType;
                if (elementType == null)
                {
                    TaskDialog.Show("BimTasksV2", "Could not resolve element type.");
                    return;
                }

                // Request placement of this type (Revit takes over placement UI)
                uidoc.PostRequestForElementTypePlacement(elementType);

                Log.Information("CreateSelectedItem: Requested placement of type '{TypeName}'", elementType.Name);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled selection
            }
            catch (Exception ex)
            {
                Log.Error(ex, "CreateSelectedItem failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
