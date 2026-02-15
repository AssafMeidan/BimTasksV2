using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Writes Uniformat assembly codes to Revit elements using shared parameters.
    /// </summary>
    public sealed class RevitUniformatWriter : IRevitUniformatWriter
    {
        private const string ParamCode = "AssemblyCode_HE";
        private const string ParamName = "AssemblyName_HE";

        public void EnsureParameters(Document doc, IEnumerable<BuiltInCategory> targets, bool bindToTypes)
        {
            var catSet = new CategorySet();
            foreach (var bic in targets)
            {
                var cat = Category.GetCategory(doc, bic);
                if (cat != null) catSet.Insert(cat);
            }

            using var group = new TransactionGroup(doc, "Ensure Uniformat Shared Parameters");
            group.Start();

            using (var t = new Transaction(doc, "Bind " + ParamCode))
            {
                t.Start();
                CreateOrBindText(doc, catSet, ParamCode, bindToTypes);
                t.Commit();
            }

            using (var t = new Transaction(doc, "Bind " + ParamName))
            {
                t.Start();
                CreateOrBindText(doc, catSet, ParamName, bindToTypes);
                t.Commit();
            }

            group.Assimilate();
        }

        public int Apply(
            UIDocument uidoc,
            string uniformatCode,
            string uniformatName,
            ApplyScope scope,
            TargetMode targetMode,
            IEnumerable<BuiltInCategory> categories)
        {
            var doc = uidoc.Document;
            IEnumerable<Element> candidates = Enumerable.Empty<Element>();

            if (scope == ApplyScope.Types)
            {
                var list = new List<Element>();
                foreach (var bic in categories)
                {
                    list.AddRange(new FilteredElementCollector(doc)
                        .OfCategory(bic)
                        .WhereElementIsElementType()
                        .ToElements());
                }
                candidates = list;
            }
            else
            {
                if (targetMode == TargetMode.Selection)
                {
                    var ids = uidoc.Selection.GetElementIds();
                    candidates = ids.Select(id => doc.GetElement(id)).Where(e => e != null)!;
                }
                else
                {
                    var list = new List<Element>();
                    foreach (var bic in categories)
                    {
                        list.AddRange(new FilteredElementCollector(doc)
                            .OfCategory(bic)
                            .WhereElementIsNotElementType()
                            .ToElements());
                    }
                    candidates = list;
                }
            }

            int updated = 0;
            using (var tx = new Transaction(doc, "Apply Uniformat"))
            {
                tx.Start();
                foreach (var e in candidates)
                {
                    bool wrote =
                        TrySet(e, ParamCode, uniformatCode) |
                        TrySet(e, ParamName, uniformatName);
                    if (wrote) updated++;
                }
                tx.Commit();
            }

            Log.Information("RevitUniformatWriter: Applied '{Code}' to {Count} elements", uniformatCode, updated);
            return updated;
        }

        #region Private Helpers

        private static void CreateOrBindText(Document doc, CategorySet catSet, string paramName, bool bindToTypes)
        {
            var app = doc.Application;
            var defFile = app.OpenSharedParameterFile()
                ?? throw new InvalidOperationException("Shared parameter file is not set.");

            var group = defFile.Groups.get_Item("BIMTasks") ?? defFile.Groups.Create("BIMTasks");

            Definition? definition = group.Definitions
                .Cast<Definition>()
                .FirstOrDefault(d => d.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase));

            if (definition == null)
            {
                var opts = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text)
                {
                    Visible = true
                };
                definition = group.Definitions.Create(opts);
            }

            ElementBinding newBinding = bindToTypes
                ? (ElementBinding)new TypeBinding(catSet)
                : new InstanceBinding(catSet);

            UpsertBinding(doc, definition, newBinding, GroupTypeId.IdentityData);
        }

        private static bool TrySet(Element e, string paramName, string value)
        {
            var p = e.LookupParameter(paramName);
            if (p is null || p.IsReadOnly) return false;
            return p.Set(value ?? string.Empty);
        }

        private static void UpsertBinding(Document doc, Definition definition, ElementBinding newBinding, ForgeTypeId groupTypeId)
        {
            var map = doc.ParameterBindings;
            var existing = GetExistingBinding(map, definition);

            if (existing.binding is null)
            {
                map.Insert(definition, newBinding, groupTypeId);
                return;
            }

            var sameKind = existing.binding.GetType() == newBinding.GetType();
            var sameCats = CategoriesEqual(existing.binding.Categories, newBinding.Categories);

            if (!sameKind || !sameCats)
                map.ReInsert(definition, newBinding, groupTypeId);
        }

        private static (ElementBinding? binding, Definition? def) GetExistingBinding(BindingMap map, Definition targetDef)
        {
            var it = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                var def = it.Key as Definition;
                if (def == null) continue;

                if (def.Name.Equals(targetDef.Name, StringComparison.OrdinalIgnoreCase))
                    return (it.Current as ElementBinding, def);
            }
            return (null, null);
        }

        private static bool CategoriesEqual(CategorySet a, CategorySet b)
        {
            var aa = a.Cast<Category>().Select(c => c.Id.Value).OrderBy(x => x).ToArray();
            var bb = b.Cast<Category>().Select(c => c.Id.Value).OrderBy(x => x).ToArray();
            if (aa.Length != bb.Length) return false;
            for (int i = 0; i < aa.Length; i++)
                if (aa[i] != bb[i]) return false;
            return true;
        }

        #endregion
    }
}
