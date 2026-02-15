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
    /// Calculates BI_EffUnitPrice and BI_ContractValue for COMP (lump sum) PayLine elements.
    ///
    /// How it works:
    /// 1. Finds all PayLines (Generic Models with BI_QtyBasis = COMP)
    /// 2. For each PayLine, sums physical elements with the same BI_BOQ_Code
    /// 3. Calculates: EffUnitPrice = ContractValue / SumPhysicalQty
    /// </summary>
    public class BOQCalcEffUnitPriceHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            Document doc = uiApp.ActiveUIDocument.Document;

            Log.Information("=== Starting BOQCalcEffUnitPriceHandler ===");

            try
            {
                // Step 1: Collect all physical pay items grouped by BOQ code
                var physicalQty = CollectPhysicalQuantities(doc);
                Log.Information("Found {Count} BOQ codes with physical quantities", physicalQty.Count);

                // Step 2: Find all PayLines
                var payLines = CollectPayLines(doc);
                Log.Information("Found {Count} PayLines to process", payLines.Count);

                if (payLines.Count == 0)
                {
                    TaskDialog.Show("BIMTasks - Calc EffUnitPrice",
                        "No PayLines found.\n\n" +
                        "PayLines are Generic Model elements with:\n" +
                        "- BI_QtyBasis = COMP\n" +
                        "- BI_AnalysisBasis set (AREA/VOLUME/LENGTH/COUNT)");
                    return;
                }

                // Step 3: Calculate and write EffUnitPrice
                var results = new List<PayLineResult>();
                var errors = new List<string>();

                using (Transaction tx = new Transaction(doc, "Calculate Effective Unit Prices"))
                {
                    tx.Start();

                    foreach (var payLine in payLines)
                    {
                        var result = ProcessPayLine(payLine, physicalQty);
                        if (result.Success)
                            results.Add(result);
                        else
                            errors.Add(result.Error);
                    }

                    tx.Commit();
                }

                ReportResults(results, errors);

                Log.Information("=== BOQCalcEffUnitPriceHandler completed. Success: {Success}, Errors: {Errors} ===",
                    results.Count, errors.Count);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "BOQCalcEffUnitPriceHandler failed");
                TaskDialog.Show("BIMTasks - Error", $"An unexpected error occurred:\n\n{ex.Message}");
            }
        }

        #region Data Collection

        /// <summary>
        /// Collects all physical (non-COMP) pay items and sums their quantities by BOQ code and basis.
        /// Returns: physicalQty[boqCode][basis] = sumOfQtyValue
        /// </summary>
        private Dictionary<string, Dictionary<string, double>> CollectPhysicalQuantities(Document doc)
        {
            var result = new Dictionary<string, Dictionary<string, double>>(StringComparer.OrdinalIgnoreCase);

            var collector = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType();

            foreach (var elem in collector)
            {
                if (elem.LookupParameter("BI_BOQ_Code") == null) continue;
                if (elem.LookupParameter("BI_IsPayItem") == null) continue;

                if (!IsPayItem(elem)) continue;

                string boqCode = GetParamString(elem, "BI_BOQ_Code");
                if (string.IsNullOrWhiteSpace(boqCode)) continue;

                string qtyBasis = GetParamString(elem, "BI_QtyBasis")?.ToUpper() ?? "";
                if (qtyBasis == "COMP") continue;
                if (!IsValidBasis(qtyBasis)) continue;

                double qtyValue = GetParamDouble(elem, "BI_QtyValue");
                if (qtyValue <= 0) continue;

                if (!result.ContainsKey(boqCode))
                    result[boqCode] = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

                if (!result[boqCode].ContainsKey(qtyBasis))
                    result[boqCode][qtyBasis] = 0;

                result[boqCode][qtyBasis] += qtyValue;
            }

            return result;
        }

        private List<Element> CollectPayLines(Document doc)
        {
            var payLines = new List<Element>();

            var collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .WhereElementIsNotElementType();

            foreach (var elem in collector)
            {
                if (elem.LookupParameter("BI_BOQ_Code") == null) continue;
                if (elem.LookupParameter("BI_AnalysisBasis") == null) continue;

                string qtyBasis = GetParamString(elem, "BI_QtyBasis")?.ToUpper() ?? "";
                if (qtyBasis != "COMP") continue;

                payLines.Add(elem);
            }

            return payLines;
        }

        #endregion

        #region PayLine Processing

        private class PayLineResult
        {
            public bool Success { get; set; }
            public string Error { get; set; }
            public ElementId ElementId { get; set; }
            public string BoqCode { get; set; }
            public double ContractValue { get; set; }
            public string AnalysisBasis { get; set; }
            public double SumQty { get; set; }
            public double EffUnitPrice { get; set; }
        }

        private PayLineResult ProcessPayLine(Element payLine,
            Dictionary<string, Dictionary<string, double>> physicalQty)
        {
            var result = new PayLineResult
            {
                ElementId = payLine.Id,
                BoqCode = GetParamString(payLine, "BI_BOQ_Code"),
                AnalysisBasis = GetParamString(payLine, "BI_AnalysisBasis")?.ToUpper() ?? ""
            };

            if (string.IsNullOrWhiteSpace(result.BoqCode))
            {
                result.Error = $"Element {payLine.Id}: Missing BI_BOQ_Code";
                return result;
            }

            if (!IsValidBasis(result.AnalysisBasis))
            {
                result.Error = $"Element {payLine.Id} ({result.BoqCode}): Invalid BI_AnalysisBasis '{result.AnalysisBasis}'";
                return result;
            }

            double unitPrice = GetParamDouble(payLine, "BI_UnitPrice");
            double qtyValue = GetParamDouble(payLine, "BI_QtyValue");
            if (qtyValue <= 0) qtyValue = 1.0;

            if (unitPrice <= 0)
            {
                result.Error = $"Element {payLine.Id} ({result.BoqCode}): Missing or zero BI_UnitPrice";
                return result;
            }

            result.ContractValue = unitPrice * qtyValue;
            SetParamDouble(payLine, "BI_ContractValue", result.ContractValue);

            if (!physicalQty.ContainsKey(result.BoqCode))
            {
                result.Success = true;
                result.EffUnitPrice = 0;
                return result;
            }

            if (!physicalQty[result.BoqCode].ContainsKey(result.AnalysisBasis))
            {
                result.Success = true;
                result.EffUnitPrice = 0;
                return result;
            }

            result.SumQty = physicalQty[result.BoqCode][result.AnalysisBasis];

            if (result.SumQty <= 0)
            {
                result.Success = true;
                result.EffUnitPrice = 0;
                return result;
            }

            result.EffUnitPrice = result.ContractValue / result.SumQty;
            SetParamDouble(payLine, "BI_EffUnitPrice", result.EffUnitPrice);

            result.Success = true;
            return result;
        }

        #endregion

        #region Parameter Helpers

        private bool IsPayItem(Element elem)
        {
            var param = elem.LookupParameter("BI_IsPayItem");
            if (param == null || param.StorageType != StorageType.Integer) return false;
            return param.AsInteger() == 1;
        }

        private string GetParamString(Element elem, string paramName)
        {
            var param = elem.LookupParameter(paramName);
            if (param == null || param.StorageType != StorageType.String) return null;
            return param.AsString()?.Trim();
        }

        private double GetParamDouble(Element elem, string paramName)
        {
            var param = elem.LookupParameter(paramName);
            if (param == null || param.StorageType != StorageType.Double) return 0;
            return param.AsDouble();
        }

        private void SetParamDouble(Element elem, string paramName, double value)
        {
            var param = elem.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly && param.StorageType == StorageType.Double)
                param.Set(value);
        }

        private bool IsValidBasis(string basis)
        {
            return basis == "AREA" || basis == "VOLUME" || basis == "LENGTH" || basis == "COUNT";
        }

        #endregion

        #region Reporting

        private void ReportResults(List<PayLineResult> results, List<string> errors)
        {
            string msg = $"Processed {results.Count + errors.Count} PayLines:\n\n";

            if (results.Any())
            {
                msg += $"Calculated: {results.Count}\n\n";

                foreach (var r in results.Take(5))
                {
                    string unit = GetUnitLabel(r.AnalysisBasis);
                    if (r.EffUnitPrice > 0)
                        msg += $"- {r.BoqCode}: {r.ContractValue:N0} ILS / {r.SumQty:N1} {unit} = {r.EffUnitPrice:N0} ILS/{unit}\n";
                    else
                        msg += $"- {r.BoqCode}: {r.ContractValue:N0} ILS (no physical qty)\n";
                }

                if (results.Count > 5)
                    msg += $"... and {results.Count - 5} more\n";
            }

            if (errors.Any())
            {
                msg += $"\nErrors: {errors.Count}\n";
                foreach (var e in errors.Take(5))
                    msg += $"- {e}\n";
                if (errors.Count > 5)
                    msg += $"... and {errors.Count - 5} more\n";
            }

            TaskDialog.Show("BIMTasks - Calc EffUnitPrice", msg);
        }

        private string GetUnitLabel(string basis)
        {
            switch (basis?.ToUpper())
            {
                case "AREA": return "m2";
                case "VOLUME": return "m3";
                case "LENGTH": return "m";
                case "COUNT": return "unit";
                default: return "?";
            }
        }

        #endregion
    }
}
