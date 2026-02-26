using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.ViewModels;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Shows the CopyCategoryFromLinkView dialog for copying categories from linked models.
    /// After dialog closes, copies elements with batch fallback on failure.
    /// </summary>
    public class CopyCategoryFromLinkHandler : ICommandHandler
    {
        private const int BatchSize = 20;
        private const int MaxBatchFailures = 5;

        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;

                // Update context service
                var contextService = container.Resolve<Services.IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uidoc;

                // Resolve dialog and get ViewModel
                var dialog = container.Resolve<Views.CopyCategoryFromLinkView>();
                var viewModel = dialog.DataContext as CopyCategoryFromLinkViewModel;

                if (viewModel == null)
                {
                    Log.Error("CopyCategoryFromLink: ViewModel is null");
                    TaskDialog.Show("Error", "Failed to initialize dialog.");
                    return;
                }

                // Initialize ViewModel with current document
                viewModel.Initialize(doc);

                // Wire up close/cancel actions
                viewModel.CloseAction = () =>
                {
                    dialog.DialogResult = true;
                    dialog.Close();
                };

                viewModel.CancelAction = () =>
                {
                    dialog.DialogResult = false;
                    dialog.Close();
                };

                // Show dialog
                bool? dialogResult = dialog.ShowDialog();

                if (dialogResult != true)
                {
                    Log.Information("CopyCategoryFromLink: User cancelled");
                    return;
                }

                // Get selections
                var selectedLink = viewModel.SelectedLink;
                var selectedCategory = viewModel.SelectedCategory;
                var selectedFamily = viewModel.SelectedFamily;

                if (selectedLink == null || selectedCategory == null)
                {
                    Log.Warning("CopyCategoryFromLink: No link or category selected");
                    return;
                }

                Document? linkDoc = selectedLink.GetLinkDocument();
                if (linkDoc == null)
                {
                    TaskDialog.Show("Error", "The selected link is no longer accessible.");
                    return;
                }

                Transform linkTransform = selectedLink.GetTotalTransform();

                // Get element IDs to copy
                List<ElementId> idsToCopy = viewModel.GetElementIdsToCopy();

                Log.Information("CopyCategoryFromLink: Copying {Count} elements from '{Link}'",
                    idsToCopy.Count, selectedLink.Name);

                if (idsToCopy.Count == 0)
                {
                    TaskDialog.Show("Info", "No elements found matching the selected criteria.");
                    return;
                }

                // Execute copy with batch fallback
                var result = ExecuteCopy(doc, linkDoc, idsToCopy, linkTransform, uidoc);

                // Show results
                ShowResults(result, idsToCopy.Count, selectedLink.Name,
                    selectedCategory.Name, selectedFamily?.Name);

                Log.Information("CopyCategoryFromLink: Completed - Copied {Copied}/{Total}, {Failed} failed",
                    result.CopiedCount, idsToCopy.Count, result.FailedCount);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "CopyCategoryFromLink failed");
                TaskDialog.Show("Error", $"Error: {ex.Message}");
            }
        }

        private CopyResult ExecuteCopy(Document doc, Document linkDoc,
            List<ElementId> idsToCopy, Transform linkTransform, UIDocument uidoc)
        {
            var copyOptions = new CopyPasteOptions();
            copyOptions.SetDuplicateTypeNamesHandler(new UseDestinationTypesHandler());

            var allNewIds = new List<ElementId>();

            // Try batch copy of all elements first (fast path)
            try
            {
                using (var t = new Transaction(doc, "Copy Elements from Link"))
                {
                    t.Start();

                    var newIds = ElementTransformUtils.CopyElements(
                        linkDoc, idsToCopy, doc, linkTransform, copyOptions);

                    t.Commit();

                    if (newIds != null)
                        allNewIds.AddRange(newIds);
                }

                Log.Information("CopyCategoryFromLink: Batch copy succeeded: {Count} elements", allNewIds.Count);

                if (allNewIds.Count > 0)
                    uidoc.Selection.SetElementIds(allNewIds);

                return new CopyResult { CopiedCount = allNewIds.Count };
            }
            catch (Exception ex)
            {
                Log.Warning("CopyCategoryFromLink: Batch copy failed, falling back to chunked copy: {Error}",
                    ex.Message);
            }

            // Fallback: copy in chunks of BatchSize, skip failed chunks
            var remaining = new List<ElementId>(idsToCopy);
            var failedIds = new List<ElementId>();
            int batchFailures = 0;

            while (remaining.Count > 0 && batchFailures < MaxBatchFailures)
            {
                var chunk = remaining.Take(BatchSize).ToList();
                remaining = remaining.Skip(BatchSize).ToList();

                try
                {
                    using (var t = new Transaction(doc, "Copy Elements from Link (batch)"))
                    {
                        t.Start();

                        var newIds = ElementTransformUtils.CopyElements(
                            linkDoc, chunk, doc, linkTransform, copyOptions);

                        t.Commit();

                        if (newIds != null)
                            allNewIds.AddRange(newIds);
                    }

                    Log.Debug("CopyCategoryFromLink: Chunk copied: {Count} elements", chunk.Count);
                }
                catch (Exception ex)
                {
                    batchFailures++;
                    failedIds.AddRange(chunk);

                    Log.Warning("CopyCategoryFromLink: Chunk failed ({Failure}/{Max}): {Error}",
                        batchFailures, MaxBatchFailures, ex.Message);
                }
            }

            // If we hit the failure limit, log the remaining as skipped
            if (batchFailures >= MaxBatchFailures && remaining.Count > 0)
            {
                failedIds.AddRange(remaining);
                Log.Error("CopyCategoryFromLink: Stopped after {Max} batch failures. {Remaining} elements skipped.",
                    MaxBatchFailures, remaining.Count);
            }

            // Log failed element details
            if (failedIds.Count > 0)
            {
                var failedInfo = failedIds.Select(id =>
                {
                    var elem = linkDoc.GetElement(id);
                    return elem != null ? $"{elem.Name} (Id:{id.IntegerValue})" : $"Id:{id.IntegerValue}";
                });
                Log.Warning("CopyCategoryFromLink: Failed elements: {Elements}",
                    string.Join(", ", failedInfo));
            }

            if (allNewIds.Count > 0)
                uidoc.Selection.SetElementIds(allNewIds);

            return new CopyResult
            {
                CopiedCount = allNewIds.Count,
                FailedCount = failedIds.Count,
                HitFailureLimit = batchFailures >= MaxBatchFailures
            };
        }

        private void ShowResults(CopyResult result, int requestedCount,
            string linkName, string categoryName, string? familyName)
        {
            string title = result.FailedCount > 0 ? "Copy Completed with Errors" : "Copy Complete";
            string instruction = result.FailedCount > 0
                ? $"Copied {result.CopiedCount} of {requestedCount} elements ({result.FailedCount} failed)"
                : $"Successfully copied {result.CopiedCount} elements";

            var td = new TaskDialog(title)
            {
                MainInstruction = instruction,
                MainContent = $"Source: {linkName}\n" +
                              $"Category: {categoryName}\n" +
                              $"Family: {familyName ?? "<All>"}"
            };

            if (result.HitFailureLimit)
            {
                td.MainContent += $"\n\nStopped early after {MaxBatchFailures} failed batches. Check logs for details.";
            }

            if (result.CopiedCount > 0)
            {
                td.FooterText = "The copied elements are now selected.";
            }

            td.Show();
        }

        private class CopyResult
        {
            public int CopiedCount { get; set; }
            public int FailedCount { get; set; }
            public bool HitFailureLimit { get; set; }
        }

        private class UseDestinationTypesHandler : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
    }
}
