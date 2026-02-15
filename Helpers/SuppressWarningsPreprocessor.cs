using Autodesk.Revit.DB;

namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Suppresses Revit warning dialogs during batch operations such as
    /// geometry joins, wall creation, etc. Warnings are deleted and
    /// resolvable errors are auto-resolved so that transactions can proceed.
    /// </summary>
    public class SuppressWarningsPreprocessor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            foreach (var failure in failuresAccessor.GetFailureMessages())
            {
                var severity = failure.GetSeverity();

                if (severity == FailureSeverity.Warning)
                {
                    failuresAccessor.DeleteWarning(failure);
                }
                else if (severity == FailureSeverity.Error && failure.HasResolutions())
                {
                    failuresAccessor.ResolveFailure(failure);
                }
            }

            return FailureProcessingResult.Continue;
        }
    }
}
