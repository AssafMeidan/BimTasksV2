using Autodesk.Revit.DB;

namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Family load options that automatically overwrite existing families
    /// and their parameter values without prompting the user.
    /// </summary>
    public class FamilyLoadOption : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }

        public bool OnSharedFamilyFound(
            Family sharedFamily,
            bool familyInUse,
            out FamilySource source,
            out bool overwriteParameterValues)
        {
            source = default;
            overwriteParameterValues = true;
            return true;
        }
    }
}
