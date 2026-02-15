using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Creates 10 BOQ and payment tracking schedules:
    /// - BI_MASTER_BOQ, BI_PAYMENT_STATUS, BI_PAYMENT_COMPONENTS, BI_ANALYSIS_UNIT_COST
    /// - BI_QA_MISSING_BOQ, BI_QA_MISSING_QTY
    /// - BI_STAGE_{FOUNDATIONS,FRAME,BUILDING,FINISHES}
    /// </summary>
    public class BOQCreateSchedulesHandler : ICommandHandler
    {
        private static readonly string[] WorkStages = { "FOUNDATIONS", "FRAME", "BUILDING", "FINISHES" };

        private Document _doc;

        public void Execute(UIApplication uiApp)
        {
            _doc = uiApp.ActiveUIDocument.Document;

            Log.Information("=== Starting BOQCreateSchedulesHandler ===");

            try
            {
                var results = new List<string>();
                var errors = new List<string>();

                using (Transaction tx = new Transaction(_doc, "Create BOQ Schedules"))
                {
                    tx.Start();

                    TryCreateSchedule(() => CreateMasterBOQ(), "BI_MASTER_BOQ", results, errors);
                    TryCreateSchedule(() => CreatePaymentStatus(), "BI_PAYMENT_STATUS", results, errors);
                    TryCreateSchedule(() => CreatePaymentComponents(), "BI_PAYMENT_COMPONENTS", results, errors);
                    TryCreateSchedule(() => CreateAnalysisUnitCost(), "BI_ANALYSIS_UNIT_COST", results, errors);
                    TryCreateSchedule(() => CreateQAMissingBOQ(), "BI_QA_MISSING_BOQ", results, errors);
                    TryCreateSchedule(() => CreateQAMissingQty(), "BI_QA_MISSING_QTY", results, errors);

                    foreach (var stage in WorkStages)
                    {
                        string name = $"BI_STAGE_{stage}";
                        TryCreateSchedule(() => CreateStageSchedule(stage), name, results, errors);
                    }

                    tx.Commit();
                }

                ReportResults(results, errors);

                Log.Information("=== BOQCreateSchedulesHandler completed. Created: {Success}, Errors: {Errors} ===",
                    results.Count, errors.Count);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "BOQCreateSchedulesHandler failed");
                TaskDialog.Show("BIMTasks - Error", $"An unexpected error occurred:\n\n{ex.Message}");
            }
        }

        #region Schedule Creation Helpers

        private void TryCreateSchedule(Func<ViewSchedule> createFunc, string name,
            List<string> results, List<string> errors)
        {
            try
            {
                createFunc();
                results.Add($"{name}: Created");
            }
            catch (Exception ex)
            {
                errors.Add($"{name}: {ex.Message}");
                Log.Warning(ex, "Failed to create schedule: {Name}", name);
            }
        }

        private string GetUniqueName(string baseName)
        {
            var existing = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(s => s.Name == baseName);

            if (existing == null) return baseName;

            for (int i = 2; i < 100; i++)
            {
                string newName = $"{baseName}_v{i}";
                existing = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewSchedule))
                    .Cast<ViewSchedule>()
                    .FirstOrDefault(s => s.Name == newName);

                if (existing == null) return newName;
            }

            throw new Exception($"Too many versions of {baseName}");
        }

        private SchedulableField FindField(ScheduleDefinition def, string paramName)
        {
            foreach (var sf in def.GetSchedulableFields())
            {
                try
                {
                    if (sf.GetName(_doc) == paramName)
                        return sf;
                }
                catch { }
            }
            return null;
        }

        private ScheduleField AddField(ScheduleDefinition def, string paramName, string heading = null)
        {
            var sf = FindField(def, paramName);
            if (sf == null) return null;

            var field = def.AddField(sf);
            if (heading != null)
                field.ColumnHeading = heading;

            return field;
        }

        private ScheduleField AddFieldWithTotal(ScheduleDefinition def, string paramName, string heading = null)
        {
            var field = AddField(def, paramName, heading);
            if (field != null)
            {
                try { field.DisplayType = ScheduleFieldDisplayType.Totals; }
                catch { }
            }
            return field;
        }

        private void AddSortGroup(ScheduleDefinition def, ScheduleFieldId fieldId,
            bool showHeader = true, bool showFooter = true, bool showBlank = true)
        {
            try
            {
                var sgf = new ScheduleSortGroupField(fieldId, ScheduleSortOrder.Ascending);
                sgf.ShowHeader = showHeader;
                sgf.ShowFooter = showFooter;
                sgf.ShowBlankLine = showBlank;
                def.AddSortGroupField(sgf);
            }
            catch { }
        }

        private void AddFilterHasValue(ScheduleDefinition def, ScheduleFieldId fieldId)
        {
            try { def.AddFilter(new ScheduleFilter(fieldId, ScheduleFilterType.HasValue)); }
            catch { }
        }

        private void AddFilterNoValue(ScheduleDefinition def, ScheduleFieldId fieldId)
        {
            try { def.AddFilter(new ScheduleFilter(fieldId, ScheduleFilterType.HasNoValue)); }
            catch { }
        }

        private void AddFilterEquals(ScheduleDefinition def, ScheduleFieldId fieldId, int intValue)
        {
            try { def.AddFilter(new ScheduleFilter(fieldId, ScheduleFilterType.Equal, intValue)); }
            catch { }
        }

        private void AddFilterEquals(ScheduleDefinition def, ScheduleFieldId fieldId, string strValue)
        {
            try { def.AddFilter(new ScheduleFilter(fieldId, ScheduleFilterType.Equal, strValue)); }
            catch { }
        }

        #endregion

        #region Schedule Definitions

        private ViewSchedule CreateMasterBOQ()
        {
            string name = GetUniqueName("BI_MASTER_BOQ");
            var schedule = ViewSchedule.CreateSchedule(_doc, ElementId.InvalidElementId);
            schedule.Name = name;
            var def = schedule.Definition;

            var f1 = AddField(def, "BI_WorkStage", "Work Stage");
            var f2 = AddField(def, "BI_Zone", "Zone");
            var f3 = AddField(def, "BI_BOQ_Code", "BOQ Code");
            AddField(def, "Category", "Category");
            AddField(def, "Family and Type", "Family and Type");
            AddField(def, "BI_HostElementId", "Host Element");
            AddField(def, "BI_QtyBasis", "Qty Basis");
            AddFieldWithTotal(def, "BI_QtyValue", "Quantity");
            AddField(def, "BI_UnitPrice", "Unit Price");

            if (f1 != null) AddSortGroup(def, f1.FieldId, true, true, true);
            if (f2 != null) AddSortGroup(def, f2.FieldId, true, true, true);
            if (f3 != null) AddSortGroup(def, f3.FieldId, true, true, false);

            var fPay = AddField(def, "BI_IsPayItem", "Pay Item");
            if (fPay != null) { fPay.IsHidden = true; AddFilterEquals(def, fPay.FieldId, 1); }

            def.IsItemized = false;
            def.ShowGrandTotal = true;
            def.ShowGrandTotalTitle = true;
            def.ShowGrandTotalCount = true;

            return schedule;
        }

        private ViewSchedule CreatePaymentStatus()
        {
            string name = GetUniqueName("BI_PAYMENT_STATUS");
            var schedule = ViewSchedule.CreateSchedule(_doc, ElementId.InvalidElementId);
            schedule.Name = name;
            var def = schedule.Definition;

            var f1 = AddField(def, "BI_WorkStage", "Work Stage");
            var f2 = AddField(def, "BI_Zone", "Zone");
            var f3 = AddField(def, "BI_BOQ_Code", "BOQ Code");
            AddField(def, "BI_HostElementId", "Host Element");
            AddFieldWithTotal(def, "BI_QtyValue", "BOQ Qty");
            AddFieldWithTotal(def, "BI_ExecPct_ToDate", "Exec %");
            AddField(def, "BI_PaidPct_ToDate", "Paid %");
            AddField(def, "BI_UnitPrice", "Unit Price");

            if (f1 != null) AddSortGroup(def, f1.FieldId, true, true, true);
            if (f2 != null) AddSortGroup(def, f2.FieldId, true, true, true);
            if (f3 != null) AddSortGroup(def, f3.FieldId, true, true, false);

            var fPay = AddField(def, "BI_IsPayItem", "Pay Item");
            if (fPay != null) { fPay.IsHidden = true; AddFilterEquals(def, fPay.FieldId, 1); }

            def.IsItemized = false;
            def.ShowGrandTotal = true;
            def.ShowGrandTotalTitle = true;

            return schedule;
        }

        private ViewSchedule CreatePaymentComponents()
        {
            string name = GetUniqueName("BI_PAYMENT_COMPONENTS");
            var schedule = ViewSchedule.CreateSchedule(_doc, new ElementId(BuiltInCategory.OST_GenericModel));
            schedule.Name = name;
            var def = schedule.Definition;

            var f1 = AddField(def, "BI_SourceElementId", "Source Element");
            var f2 = AddField(def, "BI_ComponentRole", "Component Role");
            AddField(def, "BI_BOQ_Code", "BOQ Code");
            AddField(def, "BI_Zone", "Zone");
            AddFieldWithTotal(def, "BI_QtyValue", "Quantity");
            AddField(def, "BI_ExecPct_ToDate", "Exec %");
            AddField(def, "BI_UnitPrice", "Unit Price");

            if (f1 != null) AddSortGroup(def, f1.FieldId, true, true, true);
            if (f2 != null) AddSortGroup(def, f2.FieldId, true, false, false);

            if (f1 != null) AddFilterHasValue(def, f1.FieldId);

            var fPay = AddField(def, "BI_IsPayItem", "Pay Item");
            if (fPay != null) { fPay.IsHidden = true; AddFilterEquals(def, fPay.FieldId, 1); }

            def.IsItemized = false;
            def.ShowGrandTotal = true;

            return schedule;
        }

        private ViewSchedule CreateAnalysisUnitCost()
        {
            string name = GetUniqueName("BI_ANALYSIS_UNIT_COST");
            var schedule = ViewSchedule.CreateSchedule(_doc, new ElementId(BuiltInCategory.OST_GenericModel));
            schedule.Name = name;
            var def = schedule.Definition;

            var f1 = AddField(def, "BI_WorkStage", "Work Stage");
            var f2 = AddField(def, "BI_Zone", "Zone");
            var f3 = AddField(def, "BI_BOQ_Code", "BOQ Code");
            AddField(def, "Family and Type", "Description");
            AddFieldWithTotal(def, "BI_ContractValue", "Contract Value");
            AddField(def, "BI_AnalysisBasis", "Analysis Basis");
            AddField(def, "BI_EffUnitPrice", "Eff Unit Price");

            if (f1 != null) AddSortGroup(def, f1.FieldId, true, true, true);
            if (f2 != null) AddSortGroup(def, f2.FieldId, true, true, true);
            if (f3 != null) AddSortGroup(def, f3.FieldId, true, false, false);

            var fBasis = AddField(def, "BI_QtyBasis", "Qty Basis");
            if (fBasis != null) { fBasis.IsHidden = true; AddFilterEquals(def, fBasis.FieldId, "COMP"); }

            var fPay = AddField(def, "BI_IsPayItem", "Pay Item");
            if (fPay != null) { fPay.IsHidden = true; AddFilterEquals(def, fPay.FieldId, 1); }

            def.IsItemized = true;
            def.ShowGrandTotal = true;

            return schedule;
        }

        private ViewSchedule CreateQAMissingBOQ()
        {
            string name = GetUniqueName("BI_QA_MISSING_BOQ");
            var schedule = ViewSchedule.CreateSchedule(_doc, ElementId.InvalidElementId);
            schedule.Name = name;
            var def = schedule.Definition;

            var f1 = AddField(def, "Category", "Category");
            AddField(def, "Family and Type", "Family and Type");
            AddField(def, "Level", "Level");
            var f4 = AddField(def, "BI_BOQ_Code", "BOQ Code");
            AddField(def, "BI_Zone", "Zone");
            AddField(def, "BI_WorkStage", "Work Stage");

            if (f1 != null) AddSortGroup(def, f1.FieldId, true, false, true);

            var fPay = AddField(def, "BI_IsPayItem", "Pay Item");
            if (fPay != null) { fPay.IsHidden = true; AddFilterEquals(def, fPay.FieldId, 1); }

            if (f4 != null) AddFilterNoValue(def, f4.FieldId);

            def.IsItemized = true;

            return schedule;
        }

        private ViewSchedule CreateQAMissingQty()
        {
            string name = GetUniqueName("BI_QA_MISSING_QTY");
            var schedule = ViewSchedule.CreateSchedule(_doc, ElementId.InvalidElementId);
            schedule.Name = name;
            var def = schedule.Definition;

            var f1 = AddField(def, "Category", "Category");
            AddField(def, "Family and Type", "Family and Type");
            AddField(def, "BI_BOQ_Code", "BOQ Code");
            var f4 = AddField(def, "BI_QtyBasis", "Qty Basis");
            AddField(def, "BI_QtyValue", "Quantity");
            AddField(def, "BI_QtyOverride", "Qty Override");
            AddField(def, "BI_QtyMultiplier", "Multiplier");

            if (f1 != null) AddSortGroup(def, f1.FieldId, true, false, true);

            var fPay = AddField(def, "BI_IsPayItem", "Pay Item");
            if (fPay != null) { fPay.IsHidden = true; AddFilterEquals(def, fPay.FieldId, 1); }

            if (f4 != null) AddFilterNoValue(def, f4.FieldId);

            def.IsItemized = true;

            return schedule;
        }

        private ViewSchedule CreateStageSchedule(string stageName)
        {
            string name = GetUniqueName($"BI_STAGE_{stageName}");
            var schedule = ViewSchedule.CreateSchedule(_doc, ElementId.InvalidElementId);
            schedule.Name = name;
            var def = schedule.Definition;

            var f1 = AddField(def, "BI_WorkStage", "Work Stage");
            if (f1 != null) f1.IsHidden = true;

            var f2 = AddField(def, "BI_Zone", "Zone");
            var f3 = AddField(def, "BI_BOQ_Code", "BOQ Code");
            AddField(def, "Category", "Category");
            AddField(def, "Family and Type", "Family and Type");
            AddField(def, "BI_HostElementId", "Host Element");
            AddFieldWithTotal(def, "BI_QtyValue", "Quantity");
            AddField(def, "BI_UnitPrice", "Unit Price");
            AddField(def, "BI_ExecPct_ToDate", "Exec %");

            if (f2 != null) AddSortGroup(def, f2.FieldId, true, true, true);
            if (f3 != null) AddSortGroup(def, f3.FieldId, true, true, false);

            var fPay = AddField(def, "BI_IsPayItem", "Pay Item");
            if (fPay != null) { fPay.IsHidden = true; AddFilterEquals(def, fPay.FieldId, 1); }

            if (f1 != null) AddFilterEquals(def, f1.FieldId, stageName);

            def.IsItemized = false;
            def.ShowGrandTotal = true;
            def.ShowGrandTotalTitle = true;

            return schedule;
        }

        #endregion

        #region Reporting

        private void ReportResults(List<string> results, List<string> errors)
        {
            string msg;

            if (errors.Count == 0)
            {
                msg = $"Success!\n\nCreated {results.Count} schedules:\n\n" +
                      string.Join("\n", results.Select(r => $"- {r}"));
            }
            else
            {
                msg = $"Created {results.Count} schedules, {errors.Count} errors.\n\n";

                if (results.Any())
                {
                    msg += "Created:\n";
                    foreach (var r in results.Take(5))
                        msg += $"- {r}\n";
                    if (results.Count > 5)
                        msg += $"... and {results.Count - 5} more\n";
                }

                if (errors.Any())
                {
                    msg += "\nErrors:\n";
                    foreach (var e in errors)
                        msg += $"- {e}\n";
                }
            }

            TaskDialog.Show("BIMTasks - Create Schedules", msg);
        }

        #endregion
    }
}
