using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers;
using BimTasksV2.Helpers.WallSplitter;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Reads split results from the JSON file saved by SplitWallHandler,
    /// then trims/extends replacement walls at corners.
    /// This runs as a completely separate command so Revit processes
    /// all pending warnings between splitting and trimming.
    /// </summary>
    public class TrimWallCornersHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var doc = uiApp.ActiveUIDocument.Document;

            try
            {
                var path = SplitWallHandler.SplitResultsPath;
                if (!File.Exists(path))
                {
                    TaskDialog.Show("BimTasksV2",
                        "No split results found. Split walls first, then click Trim Corners.");
                    return;
                }

                // Read and parse the JSON file
                var json = File.ReadAllText(path);
                var data = JsonSerializer.Deserialize<List<SplitResultData>>(json);
                if (data == null || data.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", "Split results file is empty.");
                    return;
                }

                // Rebuild SplitResult objects from the stored IDs
                var splitResults = new List<SplitResult>();
                foreach (var group in data)
                {
                    var sr = new SplitResult { Success = true };
                    foreach (var rep in group.Replacements)
                    {
                        var wallId = new ElementId(rep.WallId);
                        var wall = doc.GetElement(wallId) as Wall;
                        if (wall != null && wall.IsValidObject)
                        {
                            sr.ReplacementIds.Add((wallId, new LayerInfo { Index = rep.LayerIndex }));
                        }
                    }
                    if (sr.ReplacementIds.Count > 0)
                        splitResults.Add(sr);
                }

                if (splitResults.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", "No valid replacement walls found in the document.");
                    return;
                }

                using var tx = new Transaction(doc, "Trim Wall Corners");
                var failOpts = tx.GetFailureHandlingOptions();
                failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                tx.SetFailureHandlingOptions(failOpts);
                tx.Start();

                WallJoinReplicator.CrossJoinReplacements(doc, splitResults);

                tx.Commit();

                int wallCount = splitResults.Sum(r => r.ReplacementIds.Count);
                TaskDialog.Show("BimTasksV2",
                    $"Corner trimming complete for {splitResults.Count} wall groups ({wallCount} replacement walls).");
                Log.Information("[TrimWallCorners] Trimmed corners for {Groups} groups, {Walls} walls",
                    splitResults.Count, wallCount);

                // Delete the JSON file after successful trim
                try { File.Delete(path); }
                catch { /* non-critical */ }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[TrimWallCorners] Failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
