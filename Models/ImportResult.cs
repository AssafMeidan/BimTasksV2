using System.Collections.Generic;

namespace BimTasksV2.Models
{
    /// <summary>
    /// DTO for schedule import results.
    /// </summary>
    public class ImportResult
    {
        public string ScheduleName { get; set; } = string.Empty;
        public int RowsImported { get; set; }
        public int RowsFailed { get; set; }

        /// <summary>Updated cell count (used by ScheduleExcelRoundtripService).</summary>
        public int UpdatedCells { get; set; }

        /// <summary>Skipped (unchanged) cell count.</summary>
        public int SkippedCells { get; set; }

        /// <summary>Failed cell count.</summary>
        public int FailedCells { get; set; }

        /// <summary>Number of distinct elements affected.</summary>
        public int AffectedElements { get; set; }

        /// <summary>Human-readable error/failure messages.</summary>
        public List<string> Errors { get; set; } = new List<string>();

        /// <summary>Alias for Errors (backward compat with OldApp ImportResult).</summary>
        public List<string> Failures => Errors;
    }
}
