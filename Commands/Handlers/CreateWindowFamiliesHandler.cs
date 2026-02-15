using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Models;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Shows the CreateWindowFamiliesDialog for creating new window family types.
    /// The transaction wraps the dialog so user can OK to commit or Cancel to roll back.
    /// Loads window families from the document and passes them to the dialog.
    /// </summary>
    public class CreateWindowFamiliesHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // Update context service
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var contextService = container.Resolve<Services.IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uidoc;

                using (var tx = new Transaction(doc, "Create Window Family Types"))
                {
                    tx.Start();

                    // Collect window families from the document
                    var windowFamilies = BuildWindowFamilyList(doc);

                    // Create dialog with family data
                    var dialog = new Views.CreateWindowFamiliesDialog(windowFamilies);
                    bool? dialogResult = dialog.ShowDialog();

                    if (dialogResult == true)
                    {
                        tx.Commit();
                        Log.Information("CreateWindowFamilies: Transaction committed");
                    }
                    else
                    {
                        tx.RollBack();
                        Log.Information("CreateWindowFamilies: Transaction rolled back (user cancelled)");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "CreateWindowFamilies failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }

        private List<WindowFamilyModel> BuildWindowFamilyList(Document doc)
        {
            var result = new List<WindowFamilyModel>();

            var families = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.FamilyCategoryId.Value == (long)BuiltInCategory.OST_Windows)
                .OrderBy(f => f.Name);

            foreach (var family in families)
            {
                var model = new WindowFamilyModel
                {
                    Name = family.Name,
                    Family = family
                };

                // Load symbols (types) for this family
                foreach (var symbolId in family.GetFamilySymbolIds())
                {
                    if (doc.GetElement(symbolId) is FamilySymbol symbol)
                    {
                        double width = 0;
                        double height = 0;
                        double sillHeight = 0;

                        var widthParam = symbol.get_Parameter(BuiltInParameter.WINDOW_WIDTH)
                                      ?? symbol.LookupParameter("Width");
                        var heightParam = symbol.get_Parameter(BuiltInParameter.WINDOW_HEIGHT)
                                       ?? symbol.LookupParameter("Height");
                        var sillParam = symbol.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
                                     ?? symbol.LookupParameter("Default Sill Height");

                        if (widthParam != null && widthParam.HasValue)
                            width = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), UnitTypeId.Centimeters);
                        if (heightParam != null && heightParam.HasValue)
                            height = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), UnitTypeId.Centimeters);
                        if (sillParam != null && sillParam.HasValue)
                            sillHeight = UnitUtils.ConvertFromInternalUnits(sillParam.AsDouble(), UnitTypeId.Centimeters);

                        model.Symbols.Add(new WindowSymbolModel
                        {
                            Name = symbol.Name,
                            Width = width,
                            Height = height,
                            DefaultSillHeight = sillHeight,
                            SymbolId = symbolId
                        });
                    }
                }

                result.Add(model);
            }

            return result;
        }
    }
}
