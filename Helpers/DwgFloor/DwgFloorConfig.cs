namespace BimTasksV2.Helpers.DwgFloor
{
    public static class DwgFloorConfig
    {
        public const double ClosureTolerance = 0.01;       // feet (~3mm)
        public const double MinAreaSqm = 5.0;
        public const double MinAreaSqft = 53.82;
        public const double ShortCurveThreshold = 0.003;   // feet, slightly above Revit tolerance
        public const int MaxRecursionDepth = 1;
        public const int TessellationSegments = 20;
        public const double HorizontalNormalThreshold = 0.5;
        public const double ChainTolerance = 0.01;         // feet, for curve chaining proximity
        public const double DeduplicationTolerance = 0.005; // feet, for midpoint/endpoint dedup
    }
}
