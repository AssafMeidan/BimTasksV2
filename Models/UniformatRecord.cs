using System.Collections.Generic;

namespace BimTasksV2.Models
{
    /// <summary>
    /// Flat Uniformat record loaded from CSV.
    /// </summary>
    public sealed class UniformatRecord
    {
        public string RevitAssemblyCode { get; init; } = string.Empty;
        public string UniformatName { get; init; } = string.Empty;
        public string Level { get; init; } = string.Empty;
        public string BuiltInCatHint { get; init; } = string.Empty;
        public string Root { get; init; } = string.Empty;
        public int Depth { get; init; }
    }

    /// <summary>
    /// Tree node for Uniformat hierarchy display.
    /// </summary>
    public sealed class UniformatNode
    {
        public string Code { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string Level { get; init; } = string.Empty;
        public UniformatNode? Parent { get; set; }
        public List<UniformatNode> Children { get; } = new();
        public bool IsExpanded { get; set; }
    }
}
