using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Sets the phase of selected windows and doors to match their host wall's phase.
    /// Skips elements that cannot be updated (e.g., doors/windows in curtain walls).
    /// </summary>
    public class SetSelectedWindowAndDoorPhaseToWallPhaseHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Windows,
                BuiltInCategory.OST_Doors
            };

            var collector = new FilteredElementCollector(doc, uidoc.Selection.GetElementIds())
                .WherePasses(new ElementMulticategoryFilter(categories));

            int updatedWindowsCount = 0;
            int updatedDoorsCount = 0;
            var failedElements = new List<ElementId>();

            using (Transaction t = new Transaction(doc, "SetWindowAndDoorPhaseToWall"))
            {
                t.Start();

                foreach (Element hostedElement in collector)
                {
                    try
                    {
                        Parameter hostPhaseParam = GetHostPhase(doc, hostedElement);

                        if (hostPhaseParam != null)
                        {
                            Parameter elementPhaseParam = hostedElement.get_Parameter(BuiltInParameter.PHASE_CREATED);
                            if (elementPhaseParam != null && !elementPhaseParam.IsReadOnly)
                            {
                                elementPhaseParam.Set(hostPhaseParam.AsElementId());

                                if (hostedElement.Category.Id.Value == (long)BuiltInCategory.OST_Windows)
                                {
                                    updatedWindowsCount++;
                                }
                                else if (hostedElement.Category.Id.Value == (long)BuiltInCategory.OST_Doors)
                                {
                                    updatedDoorsCount++;
                                }
                            }
                            else
                            {
                                failedElements.Add(hostedElement.Id);
                            }
                        }
                        else
                        {
                            failedElements.Add(hostedElement.Id);
                        }
                    }
                    catch
                    {
                        failedElements.Add(hostedElement.Id);
                    }
                }

                t.Commit();
            }

            string messageBoxText = "Phases have been updated for selected windows and doors.";

            if (updatedWindowsCount > 0)
            {
                messageBoxText += $"\nUpdated Windows: {updatedWindowsCount}";
            }

            if (updatedDoorsCount > 0)
            {
                messageBoxText += $"\nUpdated Doors: {updatedDoorsCount}";
            }

            if (failedElements.Count > 0)
            {
                messageBoxText += "\n\nSome elements could not be processed (e.g., doors/windows in curtain walls).";
            }

            TaskDialog.Show("Completed", messageBoxText);
        }

        private static Parameter GetHostPhase(Document doc, Element hostedElement)
        {
            ElementId hostId = hostedElement.get_Parameter(BuiltInParameter.HOST_ID_PARAM)?.AsElementId()
                               ?? ElementId.InvalidElementId;

            if (hostId == ElementId.InvalidElementId)
                return null;

            Element hostElement = doc.GetElement(hostId);
            if (hostElement == null)
                return null;

            Parameter hostPhaseParam = hostElement.get_Parameter(BuiltInParameter.PHASE_CREATED);
            return hostPhaseParam;
        }
    }
}
