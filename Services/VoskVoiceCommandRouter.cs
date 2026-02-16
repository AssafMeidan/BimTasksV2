using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Events;
using Prism.Events;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Maps recognized Vosk voice inputs to Revit PostableCommand or custom BimTasksV2 actions.
    /// Uses System.Text.Json for grammar generation (not Newtonsoft.Json).
    /// </summary>
    public static class VoskVoiceCommandRouter
    {
        /// <summary>
        /// Maps voice input strings to PostableCommand values.
        /// Case-insensitive lookup.
        /// </summary>
        private static readonly Dictionary<string, PostableCommand> _commandMap = new(StringComparer.OrdinalIgnoreCase)
        {
            // --- BASIC MODELING ---
            { "area", PostableCommand.Area },
            { "area boundary", PostableCommand.AreaBoundary },
            { "ceiling", PostableCommand.AutomaticCeiling },
            { "column", PostableCommand.StructuralColumn },
            { "component", PostableCommand.PlaceAComponent },
            { "door", PostableCommand.Door },
            { "floor", PostableCommand.ArchitecturalFloor },
            { "foundation", PostableCommand.Isolated },
            { "model line", PostableCommand.ModelLine },
            { "opening by face", PostableCommand.OpeningByFace },
            { "railing", PostableCommand.Railing },
            { "region", PostableCommand.FilledRegion },
            { "room", PostableCommand.Room },
            { "room separator", PostableCommand.RoomSeparator },
            { "same", PostableCommand.CreateSimilar },
            { "create similar", PostableCommand.CreateSimilar },
            { "stair", PostableCommand.Stair },
            { "space", PostableCommand.Space },
            { "space separator", PostableCommand.SpaceSeparator },
            { "wall", PostableCommand.ArchitecturalWall },
            { "window", PostableCommand.Window },
            { "zone", PostableCommand.Zone },
            { "beam", PostableCommand.Beam },
            { "brace", PostableCommand.Brace },
            { "roof", PostableCommand.RoofByFootprint },
            { "ramp", PostableCommand.Ramp },

            // --- ANNOTATION ---
            { "align dimension", PostableCommand.AlignedDimension },
            { "angular dimension", PostableCommand.AngularDimension },
            { "linear dimension", PostableCommand.LinearDimension },
            { "material tag", PostableCommand.MaterialTag },
            { "measure", PostableCommand.MeasureBetweenTwoReferences },
            { "measure between", PostableCommand.MeasureBetweenTwoReferences },
            { "measure dimension", PostableCommand.MeasureAlongAnElement },
            { "multiple category", PostableCommand.TagAllNotTagged },
            { "spot elevation", PostableCommand.SpotElevation },
            { "tag all", PostableCommand.TagAllNotTagged },
            { "tag by category", PostableCommand.TagByCategory },
            { "revision cloud", PostableCommand.RevisionCloud },
            { "text", PostableCommand.Text },

            // --- MODIFY TOOLS ---
            { "align", PostableCommand.Align },
            { "array", PostableCommand.Array },
            { "copy", PostableCommand.CopyToClipboard },
            { "copy element", PostableCommand.Copy },
            { "match type properties", PostableCommand.MatchTypeProperties },
            { "cut", PostableCommand.CutGeometry },
            { "delete", PostableCommand.Delete },
            { "group", PostableCommand.CreateGroup },
            { "join", PostableCommand.JoinGeometry },
            { "match", PostableCommand.MatchTypeProperties },
            { "mirror", PostableCommand.MirrorDrawAxis },
            { "move", PostableCommand.Move },
            { "offset", PostableCommand.Offset },
            { "paint", PostableCommand.Paint },
            { "paste to current", PostableCommand.AlignedToCurrentView },
            { "paste to levels", PostableCommand.AlignedToSelectedLevels },
            { "paste to views", PostableCommand.AlignedToSelectedViews },
            { "paste to same place", PostableCommand.AlignedToSamePlace },
            { "paste", PostableCommand.AlignedToSamePlace },
            { "remove paint", PostableCommand.RemovePaint },
            { "repeat", PostableCommand.RepeatComponent },
            { "rotate", PostableCommand.Rotate },
            { "set work plane", PostableCommand.SetWorkPlane },
            { "split", PostableCommand.SplitElement },
            { "split element", PostableCommand.SplitElement },
            { "trim", PostableCommand.TrimOrExtendToCorner },
            { "extend", PostableCommand.TrimOrExtendToCorner },
            { "trim extend", PostableCommand.TrimOrExtendToCorner },
            { "trim extend multiple", PostableCommand.TrimOrExtendMultipleElements },
            { "type properties", PostableCommand.TypeProperties },
            { "unjoin", PostableCommand.UnjoinGeometry },
            { "uncut", PostableCommand.UncutGeometry },
            { "wall joins", PostableCommand.WallJoins },
            { "edit boundaries", PostableCommand.PickToEdit },
            { "save project", PostableCommand.Save },
            { "pin", PostableCommand.Pin },
            { "unpin", PostableCommand.Unpin },
            // "select all" removed — PostableCommand.SelectAllInstances not available in Revit 2025

            // --- DATUM & HELPERS ---
            { "grid", PostableCommand.Grid },
            { "level", PostableCommand.Level },
            { "model text", PostableCommand.ModelText },
            { "pick a plane", PostableCommand.PickAPlane },
            { "reference plane", PostableCommand.ReferencePlane },
            { "show work plane", PostableCommand.ShowWorkPlane },
            { "save selection", PostableCommand.SaveSelection },
            { "show selection", PostableCommand.LoadSelection },

            // --- VIEW & GRAPHICS ---
            { "browser organization", PostableCommand.BrowserOrganization },
            { "filters", PostableCommand.Filters },
            { "project browser", PostableCommand.ProjectBrowser },
            { "properties", PostableCommand.TogglePropertiesPalette },
            { "properties browser", PostableCommand.TogglePropertiesPalette },
            { "section", PostableCommand.Section },
            { "elevation", PostableCommand.BuildingElevation },
            { "plan", PostableCommand.FloorPlan },
            { "three d", PostableCommand.Default3DView },
            { "3d", PostableCommand.Default3DView },
            { "system browser", PostableCommand.SystemBrowser },
            { "tile views", PostableCommand.TileViews },
            { "view graphics", PostableCommand.VisibilityOrGraphics },
            { "visibility graphics", PostableCommand.VisibilityOrGraphics },
            { "close hidden", PostableCommand.CloseInactiveViews },
            { "detail line", PostableCommand.DetailLine },
            { "drafting view", PostableCommand.DraftingView },
            { "callout", PostableCommand.Callout },
            { "sheet", PostableCommand.NewSheet },
            { "schedule", PostableCommand.ScheduleOrQuantities },

            // --- MANAGE (PARAMETERS, PHASES, MATERIALS, LIBRARIES) ---
            { "global parameters", PostableCommand.GlobalParameters },
            { "load autodesk family", PostableCommand.LoadAutodeskFamily },
            { "materials", PostableCommand.Materials },
            { "phases", PostableCommand.Phases },
            { "project parameters", PostableCommand.ProjectParameters },
            { "shared parameters", PostableCommand.SharedParameters },
            { "worksets", PostableCommand.Worksets },
            { "purge", PostableCommand.PurgeUnused },
            { "transfer standards", PostableCommand.TransferProjectStandards },
            { "manage links", PostableCommand.ManageLinks },
            { "load family", PostableCommand.LoadShapes },

            // --- MEP (HVAC / Piping / Electrical) ---
            { "air terminal", PostableCommand.PlaceAComponent },
            { "cable tray", PostableCommand.CableTray },
            { "cable tray fitting", PostableCommand.CableTrayFitting },
            { "conduit", PostableCommand.Conduit },
            { "conduit fitting", PostableCommand.ConduitFitting },
            { "duct", PostableCommand.Duct },
            { "duct accessory", PostableCommand.DuctAccessory },
            { "duct fitting", PostableCommand.DuctFitting },
            { "duct placeholder", PostableCommand.DuctPlaceholder },
            { "flex duct", PostableCommand.FlexDuct },
            { "lighting fixture", PostableCommand.PlaceAComponent },
            { "mechanical equipment", PostableCommand.PlaceAComponent },
            { "pipe", PostableCommand.Pipe },
            { "pipe accessory", PostableCommand.PipeAccessory },
            { "pipe fitting", PostableCommand.PipeFitting },
            { "pipe placeholder", PostableCommand.PipePlaceholder },
            { "flex pipe", PostableCommand.FlexPipe },
            { "plumbing fixture", PostableCommand.PlaceAComponent },
            { "sprinkler", PostableCommand.Sprinkler },
            { "electrical equipment", PostableCommand.PlaceAComponent },

            // --- UNDO / REDO ---
            { "redo", PostableCommand.Redo },
            { "undo", PostableCommand.Undo },

            // --- REVIT 2025 NEW COMMANDS ---
            { "toposolid", PostableCommand.Toposolid },
            { "toposolid by face", PostableCommand.ToposolidByFace },
            { "graded region", PostableCommand.GradedRegion },
            { "smooth shading", PostableCommand.ToposolidSmoothShading },
            { "coordination model", PostableCommand.CoordinationModelAutodeskDocs },
            { "coordination local", PostableCommand.CoordinationModelLocal },
            { "shared views", PostableCommand.SharedViews },
            { "canvas theme", PostableCommand.CanvasTheme },
            { "show warnings in views", PostableCommand.ShowWarningsInViews },

            // --- VIEW TOOLS ---
            { "camera", PostableCommand.Camera },
            { "render", PostableCommand.Render },
            { "thin lines", PostableCommand.ThinLines },
            { "duplicate view", PostableCommand.DuplicateView },
            { "duplicate with detailing", PostableCommand.DuplicateWithDetailing },
            { "selection box", PostableCommand.SelectionBox },
            { "scope box", PostableCommand.ScopeBox },
            { "reflected ceiling", PostableCommand.ReflectedCeilingPlan },
            { "legend", PostableCommand.Legend },
            { "hide elements", PostableCommand.HideElements },
            { "hide category", PostableCommand.HideCategory },
            { "reveal hidden", PostableCommand.ToggleRevealHiddenElementsMode },

            // --- EXPORT / IMPORT / LINK ---
            { "export pdf", PostableCommand.ExportPDF },
            { "import image", PostableCommand.ImportImage },
            { "import pdf", PostableCommand.ImportPDF },
            { "link revit", PostableCommand.LinkRevit },
            { "link cad", PostableCommand.LinkCAD },
            { "print", PostableCommand.Print },

            // --- AUTOMATION ---
            { "dynamo", PostableCommand.Dynamo },
            { "dynamo player", PostableCommand.DynamoPlayer },
            { "synchronize", PostableCommand.SynchronizeNow },
            { "review warnings", PostableCommand.ReviewWarnings },
            { "interference check", PostableCommand.RunInterferenceCheck },

            // --- ADDITIONAL MODELING ---
            { "curtain grid", PostableCommand.CurtainGrid },
            { "mullion", PostableCommand.CurtainWallMullion },
            { "model in place", PostableCommand.ModelInPlace },
            { "shaft opening", PostableCommand.ShaftOpening },
            { "wall opening", PostableCommand.WallOpening },
            { "structural floor", PostableCommand.StructuralFloor },
            { "structural wall", PostableCommand.StructuralWall },
            { "rebar", PostableCommand.StructuralRebar },
            { "split face", PostableCommand.SplitFace },
            { "linework", PostableCommand.Linework },
            { "point cloud", PostableCommand.PointCloud },

            // --- PROJECT SETTINGS ---
            { "project units", PostableCommand.ProjectUnits },
            { "project info", PostableCommand.ProjectInformation },
            { "sun settings", PostableCommand.SunSettings },

            // --- CONTEXT PANEL ---
            { "sweep wall", PostableCommand.SweepWall },
            { "reveal wall", PostableCommand.RevealWall },
            { "filter", PostableCommand.Filters },

            // --- MISC ---
            { "keyboard shortcuts", PostableCommand.KeyboardShortcuts },
            { "shortcuts", PostableCommand.KeyboardShortcuts },
            { "add autodesk family", PostableCommand.LoadAutodeskFamily },
            { "autodesk family", PostableCommand.LoadAutodeskFamily },
            { "temporary dimensions", PostableCommand.TemporaryDimensions },
            { "manage images", PostableCommand.ManageImages },
            { "design options", PostableCommand.DesignOptions },
            { "edit filters", PostableCommand.Filters },
        };

        /// <summary>
        /// Custom commands that receive UIApplication for direct API access.
        /// These trigger BimTasksV2 internal functionality or Revit view operations.
        /// </summary>
        private static readonly Dictionary<string, Action<UIApplication>> _customCommands = new(StringComparer.OrdinalIgnoreCase)
        {
            { "filter tree", ShowFilterTreeInPanel },
            { "show filter", ShowFilterTreeInPanel },
            { "search tree", ShowFilterTreeInPanel },
            { "find element", ShowFilterTreeInPanel },

            // Calculate totals in dockable panel
            { "calculate totals", ShowCalculateTotalsInPanel },
            { "calc totals", ShowCalculateTotalsInPanel },
            { "area volume", ShowCalculateTotalsInPanel },

            // Temporary hide / isolate
            { "temporary hide", TempHideElements },
            { "temporary isolate", TempIsolateElements },
            { "reset temporary", ResetTempHideIsolate },
            { "hide category temporary", TempHideCategories },
            { "isolate category temporary", TempIsolateCategories },
        };

        #region Public API

        /// <summary>
        /// Attempts to resolve a voice command string to either a PostableCommand or a custom Action.
        /// Returns true if the command was recognized.
        /// </summary>
        /// <param name="text">The recognized voice text (trimmed, lowercase).</param>
        /// <param name="cmd">The PostableCommand if matched, null otherwise.</param>
        /// <param name="customAction">The custom Action (receives UIApplication) if matched, null otherwise.</param>
        /// <returns>True if the command was recognized.</returns>
        public static bool TryGetCommand(string text, out PostableCommand? cmd, out Action<UIApplication> customAction)
        {
            cmd = null;
            customAction = null;

            if (string.IsNullOrWhiteSpace(text))
                return false;

            string trimmed = text.Trim();

            // Check PostableCommand map first
            if (_commandMap.TryGetValue(trimmed, out var postable))
            {
                cmd = postable;
                return true;
            }

            // Check custom commands
            if (_customCommands.TryGetValue(trimmed, out var action))
            {
                customAction = action;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Returns the JSON grammar string for Vosk constrained recognition.
        /// Format: ["word1", "word2", ...]
        /// Uses System.Text.Json for serialization.
        /// </summary>
        public static string GetGrammar()
        {
            var allCommands = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var key in _commandMap.Keys)
                allCommands.Add(key);

            foreach (var key in _customCommands.Keys)
                allCommands.Add(key);

            var grammar = allCommands.OrderBy(c => c).ToArray();
            return JsonSerializer.Serialize(grammar);
        }

        /// <summary>
        /// Returns all supported voice command keys (for diagnostics/UI).
        /// </summary>
        public static IEnumerable<string> GetAllSupportedCommands()
        {
            var all = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var key in _commandMap.Keys)
                all.Add(key);

            foreach (var key in _customCommands.Keys)
                all.Add(key);

            return all.OrderBy(c => c);
        }

        /// <summary>
        /// Checks if the input matches any known command (Postable or Custom).
        /// </summary>
        public static bool IsKnownCommand(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return false;

            string trimmed = input.Trim();
            return _commandMap.ContainsKey(trimmed) || _customCommands.ContainsKey(trimmed);
        }

        #endregion

        #region Custom Command Handlers

        private static void ShowFilterTreeInPanel(UIApplication uiApp)
        {
            try
            {
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var contextService = container.Resolve<IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uiApp.ActiveUIDocument;

                var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
                eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>().Publish("FilterTree");
                eventAgg.GetEvent<BimTasksEvents.ResetFilterTreeEvent>().Publish(null);

                ShowDockablePane(uiApp);
                Log.Information("Voice: opened Filter Tree in dockable panel");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to show filter tree via voice command");
            }
        }

        private static void ShowCalculateTotalsInPanel(UIApplication uiApp)
        {
            try
            {
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var contextService = container.Resolve<IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uiApp.ActiveUIDocument;

                var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
                eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>().Publish("ElementCalculation");
                eventAgg.GetEvent<BimTasksEvents.CalculateElementsEvent>().Publish(null);

                ShowDockablePane(uiApp);
                Log.Information("Voice: opened Calculate Totals in dockable panel");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to show calculate totals via voice command");
            }
        }

        private static void ShowDockablePane(UIApplication uiApp)
        {
            var paneId = BimTasksApp.DockablePaneId;
            if (paneId != null)
            {
                var pane = uiApp.GetDockablePane(paneId);
                if (pane != null && !pane.IsShown())
                    pane.Show();
            }
        }

        private static void TempHideElements(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null) return;
            var ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0) { Log.Warning("Temporary hide: no elements selected"); return; }
            uidoc.ActiveView.HideElementsTemporary(ids);
        }

        private static void TempIsolateElements(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null) return;
            var ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0) { Log.Warning("Temporary isolate: no elements selected"); return; }
            uidoc.ActiveView.IsolateElementsTemporary(ids);
        }

        private static void ResetTempHideIsolate(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null) return;
            uidoc.ActiveView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
        }

        private static void TempHideCategories(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null) return;
            var ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0) { Log.Warning("Temporary hide category: no elements selected"); return; }
            var catIds = new HashSet<ElementId>();
            foreach (var id in ids)
            {
                var elem = uidoc.Document.GetElement(id);
                if (elem?.Category != null)
                    catIds.Add(elem.Category.Id);
            }
            if (catIds.Count > 0)
                uidoc.ActiveView.HideCategoriesTemporary(catIds);
        }

        private static void TempIsolateCategories(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null) return;
            var ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0) { Log.Warning("Temporary isolate category: no elements selected"); return; }
            var catIds = new HashSet<ElementId>();
            foreach (var id in ids)
            {
                var elem = uidoc.Document.GetElement(id);
                if (elem?.Category != null)
                    catIds.Add(elem.Category.Id);
            }
            if (catIds.Count > 0)
                uidoc.ActiveView.IsolateCategoriesTemporary(catIds);
        }

        #endregion
    }
}
