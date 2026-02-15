using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using Microsoft.Win32;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Creates or updates the BIM Israel BOQ shared parameters and binds them to payable categories.
    /// Creates 19 parameters: 14 for all payable categories and 5 for Generic Model only (PayLines/COMP items).
    /// </summary>
    public class BOQSetupHandler : ICommandHandler
    {
        private const string SharedParamGroupName = "BIMIsrael_BOQ";

        private static readonly BuiltInCategory[] PayableCategories =
        {
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Roofs,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralFramingSystem,
            BuiltInCategory.OST_CurtainWallPanels,
            BuiltInCategory.OST_CurtainWallMullions,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_Stairs,
            BuiltInCategory.OST_StairsRailing,
            BuiltInCategory.OST_Railings,
            BuiltInCategory.OST_Ramps,
            BuiltInCategory.OST_Columns,
            BuiltInCategory.OST_GenericModel
        };

        private static readonly BuiltInCategory[] GenericModelOnly =
        {
            BuiltInCategory.OST_GenericModel
        };

        private class ParamDef
        {
            public string Name { get; set; }
            public ForgeTypeId SpecType { get; set; }
            public bool IsInstance { get; set; }
            public BuiltInCategory[] Categories { get; set; }
            public ForgeTypeId GroupType { get; set; }
            public string Description { get; set; }
        }

        private ParamDef[] GetParameterDefinitions()
        {
            var specString = SpecTypeId.String.Text;
            var specBool = SpecTypeId.Boolean.YesNo;
            var specNumber = SpecTypeId.Number;

            var groupData = GroupTypeId.Data;
            var groupIdentity = GroupTypeId.IdentityData;

            return new ParamDef[]
            {
                // Core BOQ parameters (all payable categories)
                new ParamDef { Name = "BI_BOQ_Code", SpecType = specString, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "BOQ/contract leaf code" },
                new ParamDef { Name = "BI_Zone", SpecType = specString, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "Zone/WBS" },
                new ParamDef { Name = "BI_WorkStage", SpecType = specString, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "FOUNDATIONS/FRAME/BUILDING/FINISHES" },
                new ParamDef { Name = "BI_CurrentStage", SpecType = specString, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "Current work stage (DRILLING/CAGE/POUR etc.)" },
                new ParamDef { Name = "BI_IsPayItem", SpecType = specBool, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "Include in QTO/BOQ/Payment schedules" },
                new ParamDef { Name = "BI_QtyBasis", SpecType = specString, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "AREA/VOLUME/LENGTH/COUNT/COMP" },
                new ParamDef { Name = "BI_QtyValue", SpecType = specNumber, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "Unified quantity value" },
                new ParamDef { Name = "BI_QtyOverride", SpecType = specNumber, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "Override quantity (rare cases)" },
                new ParamDef { Name = "BI_QtyMultiplier", SpecType = specNumber, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "Qty multiplier (default 1.0)" },
                new ParamDef { Name = "BI_UnitPrice", SpecType = specNumber, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "Unit price in ILS" },
                new ParamDef { Name = "BI_ExecPct_ToDate", SpecType = specNumber, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "0-100 execution % to date" },
                new ParamDef { Name = "BI_PaidPct_ToDate", SpecType = specNumber, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "0-100 paid % to date" },
                new ParamDef { Name = "BI_Note", SpecType = specString, IsInstance = true, Categories = PayableCategories, GroupType = groupData, Description = "Short note" },
                new ParamDef { Name = "BI_HostElementId", SpecType = specString, IsInstance = true, Categories = PayableCategories, GroupType = groupIdentity, Description = "UniqueId of host/parent element" },

                // Generic Model specific (for PayLines / COMP items)
                new ParamDef { Name = "BI_SourceElementId", SpecType = specString, IsInstance = true, Categories = GenericModelOnly, GroupType = groupIdentity, Description = "Link to source element UniqueId" },
                new ParamDef { Name = "BI_ComponentRole", SpecType = specString, IsInstance = true, Categories = GenericModelOnly, GroupType = groupIdentity, Description = "DRILLING/CAGE/CONCRETE/etc" },
                new ParamDef { Name = "BI_AnalysisBasis", SpecType = specString, IsInstance = true, Categories = GenericModelOnly, GroupType = groupData, Description = "Analysis driver: AREA/VOLUME/LENGTH/COUNT" },
                new ParamDef { Name = "BI_EffUnitPrice", SpecType = specNumber, IsInstance = true, Categories = GenericModelOnly, GroupType = groupData, Description = "Effective/implied unit price (ILS per unit)" },
                new ParamDef { Name = "BI_ContractValue", SpecType = specNumber, IsInstance = true, Categories = GenericModelOnly, GroupType = groupData, Description = "Contract value = UnitPrice x QtyValue (ILS)" },
            };
        }

        public void Execute(UIApplication uiApp)
        {
            Application app = uiApp.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            Log.Information("=== Starting BOQSetupHandler ===");

            try
            {
                // Step 1: Get or create shared parameter file
                string sharedParamPath = GetOrCreateSharedParamFile(app);
                if (string.IsNullOrEmpty(sharedParamPath))
                {
                    Log.Information("User cancelled shared parameter file selection");
                    return;
                }

                // Step 2: Open shared parameter file
                DefinitionFile defFile = app.OpenSharedParameterFile();
                if (defFile == null)
                {
                    TaskDialog.Show("Error", "Failed to open shared parameter file.");
                    return;
                }

                // Step 3: Get or create parameter group
                DefinitionGroup group = defFile.Groups.get_Item(SharedParamGroupName);
                if (group == null)
                {
                    group = defFile.Groups.Create(SharedParamGroupName);
                    Log.Information("Created shared parameter group: {GroupName}", SharedParamGroupName);
                }

                // Step 4: Create parameters and bind to categories
                var paramDefs = GetParameterDefinitions();
                var results = new List<string>();
                var errors = new List<string>();

                using (Transaction tx = new Transaction(doc, "Setup BOQ Parameters"))
                {
                    tx.Start();

                    foreach (var paramDef in paramDefs)
                    {
                        try
                        {
                            var result = CreateAndBindParameter(app, doc, group, paramDef);
                            results.Add($"{paramDef.Name}: {result}");
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"{paramDef.Name}: {ex.Message}");
                            Log.Warning(ex, "Failed to create parameter: {Name}", paramDef.Name);
                        }
                    }

                    tx.Commit();
                }

                // Report results
                ReportResults(paramDefs.Length, results, errors);

                Log.Information("=== BOQSetupHandler completed. {Success}/{Total} parameters ===",
                    paramDefs.Length - errors.Count, paramDefs.Length);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "BOQSetupHandler failed");
                TaskDialog.Show("BIMTasks - Error", $"An unexpected error occurred:\n\n{ex.Message}");
            }
        }

        private string GetOrCreateSharedParamFile(Application app)
        {
            string currentPath = app.SharedParametersFilename;
            if (!string.IsNullOrEmpty(currentPath) && File.Exists(currentPath))
            {
                var result = TaskDialog.Show("BIMTasks - Shared Parameters",
                    $"Current shared parameter file:\n{currentPath}\n\nUse this file?",
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (result == TaskDialogResult.Yes)
                    return currentPath;
            }

            var dlg = new SaveFileDialog
            {
                Title = "Select or Create Shared Parameter File",
                Filter = "Text Files (*.txt)|*.txt",
                FileName = "BIMIsrael_SharedParams.txt",
                OverwritePrompt = false
            };

            if (dlg.ShowDialog() != true)
                return null;

            string path = dlg.FileName;

            if (!File.Exists(path))
            {
                File.WriteAllText(path, "");
                Log.Information("Created new shared parameter file: {Path}", path);
            }

            app.SharedParametersFilename = path;
            return path;
        }

        private string CreateAndBindParameter(Application app, Document doc, DefinitionGroup group, ParamDef paramDef)
        {
            Definition definition = group.Definitions.get_Item(paramDef.Name);
            bool created = false;

            if (definition == null)
            {
                var options = new ExternalDefinitionCreationOptions(paramDef.Name, paramDef.SpecType)
                {
                    Visible = true,
                    UserModifiable = true,
                    Description = paramDef.Description
                };
                definition = group.Definitions.Create(options);
                created = true;
            }

            CategorySet catSet = app.Create.NewCategorySet();
            foreach (var bic in paramDef.Categories)
            {
                Category cat = doc.Settings.Categories.get_Item(bic);
                if (cat != null && cat.AllowsBoundParameters)
                {
                    catSet.Insert(cat);
                }
            }

            if (catSet.Size == 0)
                return "No valid categories";

            // Check if already bound
            Definition existingDef = FindBoundDefinition(doc, paramDef.Name);
            if (existingDef != null)
            {
                var binding = paramDef.IsInstance
                    ? (ElementBinding)app.Create.NewInstanceBinding(catSet)
                    : (ElementBinding)app.Create.NewTypeBinding(catSet);

                bool rebound = doc.ParameterBindings.ReInsert(existingDef, binding, paramDef.GroupType);
                return rebound ? "Rebound" : "Exists (rebind failed)";
            }

            var newBinding = paramDef.IsInstance
                ? (ElementBinding)app.Create.NewInstanceBinding(catSet)
                : (ElementBinding)app.Create.NewTypeBinding(catSet);

            bool inserted = doc.ParameterBindings.Insert(definition, newBinding, paramDef.GroupType);

            if (created && inserted) return "Created & Bound";
            if (inserted) return "Bound";
            return "Bind failed";
        }

        private Definition FindBoundDefinition(Document doc, string paramName)
        {
            var iterator = doc.ParameterBindings.ForwardIterator();
            while (iterator.MoveNext())
            {
                var def = iterator.Key;
                if (def != null && def.Name == paramName)
                    return def;
            }
            return null;
        }

        private void ReportResults(int total, List<string> results, List<string> errors)
        {
            string msg;

            if (errors.Count == 0)
            {
                msg = $"Success!\n\nAll {total} BOQ parameters are ready.\n\nParameters:\n" +
                      string.Join("\n", results.Take(10));
                if (results.Count > 10)
                    msg += $"\n... and {results.Count - 10} more";
            }
            else
            {
                msg = $"Completed with {total - errors.Count}/{total} parameters.\n\n";
                if (errors.Any())
                {
                    msg += "Errors:\n";
                    foreach (var e in errors.Take(5))
                        msg += $"  - {e}\n";
                }
            }

            TaskDialog.Show("BIMTasks - BOQ Setup", msg);
        }
    }
}
