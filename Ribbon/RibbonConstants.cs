namespace BimTasksV2.Ribbon
{
    /// <summary>
    /// Defines a ribbon button: internal name, display text, and full class name.
    /// </summary>
    public record ButtonDef(string Name, string Text, string ClassName, string Tooltip = "");

    /// <summary>
    /// Constants for the BimTasks ribbon tab, panel names, and button definitions.
    /// 27 ribbon buttons across 5 panels, plus 4 utility commands with no ribbon presence.
    /// </summary>
    public static class RibbonConstants
    {
        public const string TabName = "BimTasks";

        // Panel names (appear in creation order)
        public const string WallPanel = "Walls";
        public const string StructurePanel = "Structure";
        public const string SchedulePanel = "Schedules";
        public const string BOQPanel = "BOQ";
        public const string ToolPanel = "Tools";

        // =====================================================================
        // Panel 1: Walls (10 buttons)
        // =====================================================================

        public static readonly ButtonDef[] WallButtons = new[]
        {
            new ButtonDef(
                "btnSetWindowDoorPhase",
                "Sync\nPhases",
                "BimTasksV2.Commands.Proxies.SetSelectedWindowAndDoorPhaseToWallPhaseCommand",
                "Set the phase of selected windows and doors to match their host wall"),

            new ButtonDef(
                "btnCopyLinkedWalls",
                "Copy Linked\nWalls",
                "BimTasksV2.Commands.Proxies.CopyLinkedWallsCommand",
                "Copy walls from a linked Revit model into the current project"),

            new ButtonDef(
                "btnAdjustWallHeight",
                "Wall To\nFloors",
                "BimTasksV2.Commands.Proxies.AdjustWallHeightToFloorsCommand",
                "Extend walls up or down to match the floors above and below"),

            new ButtonDef(
                "btnChangeWallHeight",
                "Wall\nHeight",
                "BimTasksV2.Commands.Proxies.ChangeWallHeightCommand",
                "Set a new height for selected walls"),

            new ButtonDef(
                "btnSplitWall",
                "Split\nWall",
                "BimTasksV2.Commands.Proxies.SplitWallCommand",
                "Split a compound wall into individual layer walls"),

            new ButtonDef(
                "btnDetectOverlapping",
                "Detect\nOverlaps",
                "BimTasksV2.Commands.Proxies.DetectOverlappingWallsCommand",
                "Find duplicate walls stacked at the same location"),

            // Cladding split button members (added as SplitButton in RibbonBuilder)
            new ButtonDef(
                "btnAddChipuyToWall",
                "Wall Cladding",
                "BimTasksV2.Commands.Proxies.AddChipuyToWallCommand",
                "Split compound wall and add cladding layers to both sides"),

            new ButtonDef(
                "btnAddChipuyExternal",
                "External Cladding",
                "BimTasksV2.Commands.Proxies.AddChipuyToExternalWallCommand",
                "Split compound wall and add cladding to the exterior face"),

            new ButtonDef(
                "btnAddChipuyInternal",
                "Internal Cladding",
                "BimTasksV2.Commands.Proxies.AddChipuyToInternalWallCommand",
                "Split compound wall and add cladding to the interior face"),

            new ButtonDef(
                "btnCreateWindowFamilies",
                "Window\nFamilies",
                "BimTasksV2.Commands.Proxies.CreateWindowFamiliesCommand",
                "Generate parametric window families from specifications"),
        };

        // Indices into WallButtons for the cladding SplitButton
        public const int CladdingSplitStart = 6; // AddChipuyToWall
        public const int CladdingSplitEnd = 8;   // AddChipuyToInternal (inclusive)

        // =====================================================================
        // Panel 2: Structure (2 buttons)
        // =====================================================================

        public static readonly ButtonDef[] StructureButtons = new[]
        {
            new ButtonDef(
                "btnSetBeamsToGround",
                "Beams To\nGround",
                "BimTasksV2.Commands.Proxies.SetBeamsToGroundZeroCommand",
                "Move selected beams so their base offset is at ground zero"),

            new ButtonDef(
                "btnJoinConcreteWallsFloors",
                "Join Walls\n& Floors",
                "BimTasksV2.Commands.Proxies.JoinAllConcreteWallsAndFloorsGeometryCommand",
                "Join geometry between all concrete walls and floors for clean intersections"),

            new ButtonDef(
                "btnSplitFloor",
                "Split\nFloor",
                "BimTasksV2.Commands.Proxies.SplitFloorCommand",
                "Split a compound floor into individual layer floors"),

            new ButtonDef(
                "btnJoinColumnsFramesToWalls",
                "Join Col\n& Frames",
                "BimTasksV2.Commands.Proxies.JoinColumnsAndFramesToWallsCommand",
                "Join selected columns and structural framing to intersecting walls"),
        };

        // =====================================================================
        // Panel 3: Schedules (3 buttons)
        // =====================================================================

        public static readonly ButtonDef[] ScheduleButtons = new[]
        {
            new ButtonDef(
                "btnExportScheduleToExcel",
                "Export\nSchedule",
                "BimTasksV2.Commands.Proxies.ExportScheduleToExcelCommand",
                "Export the active schedule to a styled Excel file"),

            new ButtonDef(
                "btnEditScheduleInExcel",
                "Edit In\nExcel",
                "BimTasksV2.Commands.Proxies.EditScheduleInExcelCommand",
                "Open the active schedule in Excel for editing, then import changes back"),

            new ButtonDef(
                "btnImportKeySchedules",
                "Import Key\nSchedules",
                "BimTasksV2.Commands.Proxies.ImportKeySchedulesFromExcelCommand",
                "Import key schedule data from an Excel file"),
        };

        // =====================================================================
        // Panel 4: BOQ (5 buttons)
        // =====================================================================

        public static readonly ButtonDef[] BOQButtons = new[]
        {
            new ButtonDef(
                "btnBOQSetup",
                "BOQ\nSetup",
                "BimTasksV2.Commands.Proxies.BOQSetupCommand",
                "Configure BOQ parameters and Uniformat code mappings"),

            new ButtonDef(
                "btnBOQCreateSchedules",
                "BOQ\nSchedules",
                "BimTasksV2.Commands.Proxies.BOQCreateSchedulesCommand",
                "Generate BOQ schedule views from configured parameters"),

            new ButtonDef(
                "btnBOQAutoFillQty",
                "BOQ\nAutoFill",
                "BimTasksV2.Commands.Proxies.BOQAutoFillQtyCommand",
                "Automatically calculate and fill quantity values in BOQ schedules"),

            new ButtonDef(
                "btnBOQCalcEffUnitPrice",
                "BOQ Calc\nPrice",
                "BimTasksV2.Commands.Proxies.BOQCalcEffUnitPriceCommand",
                "Calculate effective unit prices for BOQ line items"),

            new ButtonDef(
                "btnCreateSeifeiChoze",
                "Contract\nSections",
                "BimTasksV2.Commands.Proxies.CreateSeifeiChozeCommand",
                "Generate contract section schedules from an Excel template"),
        };

        // =====================================================================
        // Panel 5: Tools (9 regular + SplitButton)
        // =====================================================================

        public static readonly ButtonDef[] ToolButtons = new[]
        {
            new ButtonDef(
                "btnShowFilterTree",
                "Filter\nTree",
                "BimTasksV2.Commands.Proxies.ShowFilterTreeWindowCommand",
                "Open hierarchical element filter \u2014 filter by category, parameter, and value"),

            new ButtonDef(
                "btnCheckWallsAreaVolume",
                "Calc Area\n& Volume",
                "BimTasksV2.Commands.Proxies.CheckSelectedWallsAreaAndVolumeCommand",
                "Calculate and display area and volume totals for selected elements"),

            new ButtonDef(
                "btnShowUniformat",
                "Uniformat\nCodes",
                "BimTasksV2.Commands.Proxies.ShowUniformatWindowCommand",
                "View and assign Uniformat classification codes to elements"),

            new ButtonDef(
                "btnCopyCategoryFromLink",
                "Copy From\nLink",
                "BimTasksV2.Commands.Proxies.CopyCategoryFromLinkCommand",
                "Copy elements of a specific category from a linked model"),

            new ButtonDef(
                "btnToggleDockablePanel",
                "Dockable\nPanel",
                "BimTasksV2.Commands.Proxies.ToggleDockablePanelCommand",
                "Show or hide the BimTasks dockable panel"),

            new ButtonDef(
                "btnToggleToolbar",
                "Floating\nToolbar",
                "BimTasksV2.Commands.Proxies.ToggleToolbarCommand",
                "Show or hide the floating quick-access toolbar"),

            new ButtonDef(
                "btnExportJSON",
                "Export\nJSON",
                "BimTasksV2.Commands.Proxies.WebAppDataExtractorCommand",
                "Export element data to JSON for web app integration"),

            new ButtonDef(
                "btnToggleVosk",
                "Voice\nControl",
                "BimTasksV2.Commands.Proxies.ToggleVoskVoiceRecognitionCommand",
                "Toggle voice command recognition on or off"),

            new ButtonDef(
                "btnPickAndCreateFloor",
                "Floor From\nDWG Pick",
                "BimTasksV2.Commands.Proxies.PickAndCreateFloorCommand",
                "Pick a closed region in a DWG and create a floor from it"),

            new ButtonDef(
                "btnImportAllDwgFloors",
                "Floors From\nDWG All",
                "BimTasksV2.Commands.Proxies.ImportAllDwgFloorsCommand",
                "Create floors from all closed regions in a DWG file"),

            new ButtonDef(
                "btnColorCodeByParam",
                "Color\nCode",
                "BimTasksV2.Commands.Proxies.ColorCodeByParameterCommand",
                "Color-code elements in the view by any parameter value with a clickable legend"),
        };

        /// <summary>
        /// SplitButton members in Tool panel: CreateSelectedItem (default) + AddClunasFromDwg.
        /// </summary>
        public static readonly ButtonDef[] ToolSplitButtons = new[]
        {
            new ButtonDef(
                "btnCreateSelectedItem",
                "Create\nSimilar",
                "BimTasksV2.Commands.Proxies.CreateSelectedItemCommand",
                "Create a new element similar to the currently selected one"),

            new ButtonDef(
                "btnAddClunasFromDwg",
                "Piles From\nDWG",
                "BimTasksV2.Commands.Proxies.AddClunasFromDwgCommand",
                "Create pile elements from point positions in a DWG file"),
        };

        // =====================================================================
        // Utility commands (no ribbon buttons, 4 total)
        // =====================================================================

        public static readonly ButtonDef[] UtilityButtons = new[]
        {
            new ButtonDef(
                "btnJoinBeamsToBlockWalls",
                "Join Walls\n& Beams",
                "BimTasksV2.Commands.Proxies.JoinSelectedBeamsToBlockWallsCommand",
                "Join geometry of selected beams to adjacent block walls"),

            new ButtonDef(
                "btnAddClunasFromDwg_Utility",
                "Piles From\nDWG",
                "BimTasksV2.Commands.Proxies.AddClunasFromDwgCommand",
                "Create pile elements from point positions in a DWG file"),

            new ButtonDef(
                "btnFamilyParamChanger",
                "Family Param\nChanger",
                "BimTasksV2.Commands.Proxies.FamilyParamChangerCommand",
                "Bulk-edit family parameter values across selected elements"),

            new ButtonDef(
                "btnTestInfrastructure",
                "Test\nInfrastructure",
                "BimTasksV2.Commands.TestInfrastructureCommand",
                "Run infrastructure diagnostics to verify event system and services"),
        };
    }
}
