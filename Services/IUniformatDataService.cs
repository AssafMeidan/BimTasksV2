using System.Collections.Generic;
using System.Threading.Tasks;
using BimTasksV2.Models;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Loads Uniformat records from CSV and builds a hierarchical tree.
    /// </summary>
    public interface IUniformatDataService
    {
        /// <summary>
        /// Loads flat Uniformat records from a CSV file asynchronously.
        /// </summary>
        Task<IReadOnlyList<UniformatRecord>> LoadFromCsvAsync(string csvPath);

        /// <summary>
        /// Returns all loaded records (requires prior LoadFromCsvAsync call).
        /// </summary>
        IReadOnlyList<UniformatRecord> GetAllRecords();

        /// <summary>
        /// Searches loaded records by code or name substring.
        /// </summary>
        IReadOnlyList<UniformatRecord> Search(string query);

        /// <summary>
        /// Builds a hierarchical tree from the flat record list.
        /// </summary>
        UniformatNode BuildTree(IReadOnlyList<UniformatRecord> flatRecords);
    }
}
