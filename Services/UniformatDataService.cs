using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using BimTasksV2.Models;

namespace BimTasksV2.Services
{
    /// <summary>
    /// CSV-based Uniformat data loader. No external dependencies.
    /// </summary>
    public sealed class UniformatDataService : IUniformatDataService
    {
        private IReadOnlyList<UniformatRecord> _records = Array.Empty<UniformatRecord>();

        public async Task<IReadOnlyList<UniformatRecord>> LoadFromCsvAsync(string csvPath)
        {
            var list = new List<UniformatRecord>();
            using var sr = new StreamReader(csvPath);
            string? header = await sr.ReadLineAsync().ConfigureAwait(false);
            if (header is null)
            {
                _records = list;
                return list;
            }

            var columns = header.Split(',');
            int idxCode = Array.IndexOf(columns, "RevitAssemblyCode");
            int idxName = Array.IndexOf(columns, "UniformatName");
            int idxLevel = Array.IndexOf(columns, "Level");
            int idxRoot = Array.IndexOf(columns, "Root");
            int idxDepth = Array.IndexOf(columns, "Depth");
            int idxHint = Array.IndexOf(columns, "BuiltInCatHint");

            string? line;
            while ((line = await sr.ReadLineAsync().ConfigureAwait(false)) != null)
            {
                var parts = SplitCsv(line);
                string code = Get(parts, idxCode);
                if (string.IsNullOrWhiteSpace(code)) continue;

                list.Add(new UniformatRecord
                {
                    RevitAssemblyCode = code.Trim(),
                    UniformatName = Get(parts, idxName),
                    Level = Get(parts, idxLevel),
                    Root = Get(parts, idxRoot),
                    Depth = int.TryParse(Get(parts, idxDepth), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out var d) ? d : 0,
                    BuiltInCatHint = Get(parts, idxHint)
                });
            }

            _records = list;
            return list;
        }

        public IReadOnlyList<UniformatRecord> GetAllRecords()
        {
            return _records;
        }

        public IReadOnlyList<UniformatRecord> Search(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
                return _records;

            string q = query.Trim();
            return _records
                .Where(r => r.RevitAssemblyCode.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                            r.UniformatName.Contains(q, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public UniformatNode BuildTree(IReadOnlyList<UniformatRecord> flatRecords)
        {
            var root = new UniformatNode { Code = "ROOT", Name = "Uniformat" };
            var lookup = flatRecords
                .GroupBy(f => f.Root)
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (var kv in lookup.OrderBy(k => k.Key))
            {
                var rootNode = new UniformatNode { Code = kv.Key, Name = kv.Key, Level = "1" };
                rootNode.Parent = root;
                root.Children.Add(rootNode);

                var items = kv.Value
                    .OrderBy(v => v.RevitAssemblyCode, StringComparer.Ordinal)
                    .ToList();

                foreach (var rec in items)
                    Insert(rootNode, rec);
            }

            return root;
        }

        #region Private Helpers

        private static void Insert(UniformatNode parent, UniformatRecord rec)
        {
            var code = rec.RevitAssemblyCode;
            var parts = code.Split('.');
            var chain = new List<string>();

            if (parts.Length == 1)
            {
                chain.Add(parts[0]);
            }
            else
            {
                chain.Add(parts[0]);
                for (int i = 1; i < parts.Length; i++)
                    chain.Add(string.Join(".", parts.Take(i + 1)));
            }

            UniformatNode current = parent;
            for (int i = 0; i < chain.Count; i++)
            {
                var key = chain[i];
                var found = current.Children.FirstOrDefault(c => c.Code == key);
                if (found == null)
                {
                    found = new UniformatNode
                    {
                        Code = key,
                        Name = i == chain.Count - 1 ? rec.UniformatName : key,
                        Level = rec.Level
                    };
                    found.Parent = current;
                    current.Children.Add(found);
                }
                current = found;
            }
        }

        private static List<string> SplitCsv(string line)
        {
            return line.Split(',').Select(s => s.Trim()).ToList();
        }

        private static string Get(List<string> parts, int idx) =>
            idx >= 0 && idx < parts.Count ? parts[idx] : string.Empty;

        #endregion
    }
}
