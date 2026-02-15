namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Represents the possible location line values for a wall.
    /// Maps to Revit's WallLocationLine built-in parameter values.
    /// </summary>
    public enum WallLocationLineValue
    {
        WallCenterline = 0,
        CoreCenterline = 1,
        FinishFaceExterior = 2,
        FinishFaceInterior = 3,
        CoreExterior = 4,
        CoreInterior = 5
    }
}
