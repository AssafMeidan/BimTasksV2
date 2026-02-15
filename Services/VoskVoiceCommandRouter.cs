using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.UI;
using BimTasksV2.Events;
using Prism.Events;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Maps recognized Vosk voice inputs to Revit PostableCommand or custom BimTasksV2 actions.
    /// Contains ~145 voice command mappings and 4 custom event-based commands.
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
        };

        /// <summary>
        /// Custom commands handled via Prism events (not PostableCommand).
        /// These trigger BimTasksV2 internal functionality.
        /// </summary>
        private static readonly Dictionary<string, Action> _customCommands = new(StringComparer.OrdinalIgnoreCase)
        {
            { "filter tree", PublishShowFilterTree },
            { "show filter", PublishShowFilterTree },
            { "search tree", PublishShowFilterTree },
            { "find element", PublishShowFilterTree },
        };

        #region Public API

        /// <summary>
        /// Attempts to resolve a voice command string to either a PostableCommand or a custom Action.
        /// Returns true if the command was recognized.
        /// </summary>
        /// <param name="text">The recognized voice text (trimmed, lowercase).</param>
        /// <param name="cmd">The PostableCommand if matched, null otherwise.</param>
        /// <param name="customAction">The custom Action if matched, null otherwise.</param>
        /// <returns>True if the command was recognized.</returns>
        public static bool TryGetCommand(string text, out PostableCommand? cmd, out Action customAction)
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

        private static void PublishShowFilterTree()
        {
            try
            {
                var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
                eventAgg.GetEvent<BimTasksEvents.ResetFilterTreeEvent>().Publish(null);
                Log.Information("Published ResetFilterTreeEvent via voice command");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to publish filter tree event from voice command");
            }
        }

        #endregion
    }
}
