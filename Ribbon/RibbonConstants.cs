namespace BimTasksV2.Ribbon
{
    /// <summary>
    /// Defines a ribbon button: internal name, display text, and full class name.
    /// </summary>
    public record ButtonDef(string Name, string Text, string ClassName);

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
        // Panel 1: Walls (9 buttons)
        // =====================================================================

        public static readonly ButtonDef[] WallButtons = new[]
        {
            new ButtonDef(
                "btnSetWindowDoorPhase",
                "Sync\nPhases",
                "BimTasksV2.Commands.Proxies.SetSelectedWindowAndDoorPhaseToWallPhaseCommand"),

            new ButtonDef(
                "btnCopyLinkedWalls",
                "Copy Linked\nWalls",
                "BimTasksV2.Commands.Proxies.CopyLinkedWallsCommand"),

            new ButtonDef(
                "btnAdjustWallHeight",
                "Wall To\nFloors",
                "BimTasksV2.Commands.Proxies.AdjustWallHeightToFloorsCommand"),

            new ButtonDef(
                "btnChangeWallHeight",
                "Wall\nHeight",
                "BimTasksV2.Commands.Proxies.ChangeWallHeightCommand"),

            new ButtonDef(
                "btnSplitWall",
                "Split\nWall",
                "BimTasksV2.Commands.Proxies.SplitWallCommand"),

            // Cladding split button members (added as SplitButton in RibbonBuilder)
            new ButtonDef(
                "btnAddChipuyToWall",
                "Wall Cladding",
                "BimTasksV2.Commands.Proxies.AddChipuyToWallCommand"),

            new ButtonDef(
                "btnAddChipuyExternal",
                "External Cladding",
                "BimTasksV2.Commands.Proxies.AddChipuyToExternalWallCommand"),

            new ButtonDef(
                "btnAddChipuyInternal",
                "Internal Cladding",
                "BimTasksV2.Commands.Proxies.AddChipuyToInternalWallCommand"),

            new ButtonDef(
                "btnCreateWindowFamilies",
                "Window\nFamilies",
                "BimTasksV2.Commands.Proxies.CreateWindowFamiliesCommand"),
        };

        // Indices into WallButtons for the cladding SplitButton
        public const int CladdingSplitStart = 5; // AddChipuyToWall
        public const int CladdingSplitEnd = 7;   // AddChipuyToInternal (inclusive)

        // =====================================================================
        // Panel 2: Structure (2 buttons)
        // =====================================================================

        public static readonly ButtonDef[] StructureButtons = new[]
        {
            new ButtonDef(
                "btnSetBeamsToGround",
                "Beams To\nGround",
                "BimTasksV2.Commands.Proxies.SetBeamsToGroundZeroCommand"),

            new ButtonDef(
                "btnJoinConcreteWallsFloors",
                "Join Walls\n& Floors",
                "BimTasksV2.Commands.Proxies.JoinAllConcreteWallsAndFloorsGeometryCommand"),
        };

        // =====================================================================
        // Panel 3: Schedules (3 buttons)
        // =====================================================================

        public static readonly ButtonDef[] ScheduleButtons = new[]
        {
            new ButtonDef(
                "btnExportScheduleToExcel",
                "Export\nSchedule",
                "BimTasksV2.Commands.Proxies.ExportScheduleToExcelCommand"),

            new ButtonDef(
                "btnEditScheduleInExcel",
                "Edit In\nExcel",
                "BimTasksV2.Commands.Proxies.EditScheduleInExcelCommand"),

            new ButtonDef(
                "btnImportKeySchedules",
                "Import Key\nSchedules",
                "BimTasksV2.Commands.Proxies.ImportKeySchedulesFromExcelCommand"),
        };

        // =====================================================================
        // Panel 4: BOQ (5 buttons)
        // =====================================================================

        public static readonly ButtonDef[] BOQButtons = new[]
        {
            new ButtonDef(
                "btnBOQSetup",
                "BOQ\nSetup",
                "BimTasksV2.Commands.Proxies.BOQSetupCommand"),

            new ButtonDef(
                "btnBOQCreateSchedules",
                "BOQ\nSchedules",
                "BimTasksV2.Commands.Proxies.BOQCreateSchedulesCommand"),

            new ButtonDef(
                "btnBOQAutoFillQty",
                "BOQ\nAutoFill",
                "BimTasksV2.Commands.Proxies.BOQAutoFillQtyCommand"),

            new ButtonDef(
                "btnBOQCalcEffUnitPrice",
                "BOQ Calc\nPrice",
                "BimTasksV2.Commands.Proxies.BOQCalcEffUnitPriceCommand"),

            new ButtonDef(
                "btnCreateSeifeiChoze",
                "Contract\nSections",
                "BimTasksV2.Commands.Proxies.CreateSeifeiChozeCommand"),
        };

        // =====================================================================
        // Panel 5: Tools (8 regular + SplitButton)
        // =====================================================================

        public static readonly ButtonDef[] ToolButtons = new[]
        {
            new ButtonDef(
                "btnShowFilterTree",
                "Filter\nTree",
                "BimTasksV2.Commands.Proxies.ShowFilterTreeWindowCommand"),

            new ButtonDef(
                "btnCheckWallsAreaVolume",
                "Calc Area\n& Volume",
                "BimTasksV2.Commands.Proxies.CheckSelectedWallsAreaAndVolumeCommand"),

            new ButtonDef(
                "btnShowUniformat",
                "Uniformat\nCodes",
                "BimTasksV2.Commands.Proxies.ShowUniformatWindowCommand"),

            new ButtonDef(
                "btnCopyCategoryFromLink",
                "Copy From\nLink",
                "BimTasksV2.Commands.Proxies.CopyCategoryFromLinkCommand"),

            new ButtonDef(
                "btnToggleDockablePanel",
                "Dockable\nPanel",
                "BimTasksV2.Commands.Proxies.ToggleDockablePanelCommand"),

            new ButtonDef(
                "btnToggleToolbar",
                "Floating\nToolbar",
                "BimTasksV2.Commands.Proxies.ToggleToolbarCommand"),

            new ButtonDef(
                "btnExportJSON",
                "Export\nJSON",
                "BimTasksV2.Commands.Proxies.WebAppDataExtractorCommand"),

            new ButtonDef(
                "btnToggleVosk",
                "Voice\nControl",
                "BimTasksV2.Commands.Proxies.ToggleVoskVoiceRecognitionCommand"),

            new ButtonDef(
                "btnPickAndCreateFloor",
                "Floor From\nDWG Pick",
                "BimTasksV2.Commands.Proxies.PickAndCreateFloorCommand"),

            new ButtonDef(
                "btnImportAllDwgFloors",
                "Floors From\nDWG All",
                "BimTasksV2.Commands.Proxies.ImportAllDwgFloorsCommand"),
        };

        /// <summary>
        /// SplitButton members in Tool panel: CreateSelectedItem (default) + AddClunasFromDwg.
        /// </summary>
        public static readonly ButtonDef[] ToolSplitButtons = new[]
        {
            new ButtonDef(
                "btnCreateSelectedItem",
                "Create\nSimilar",
                "BimTasksV2.Commands.Proxies.CreateSelectedItemCommand"),

            new ButtonDef(
                "btnAddClunasFromDwg",
                "Piles From\nDWG",
                "BimTasksV2.Commands.Proxies.AddClunasFromDwgCommand"),
        };

        // =====================================================================
        // Utility commands (no ribbon buttons, 4 total)
        // =====================================================================

        public static readonly ButtonDef[] UtilityButtons = new[]
        {
            new ButtonDef(
                "btnJoinBeamsToBlockWalls",
                "Join Walls\n& Beams",
                "BimTasksV2.Commands.Proxies.JoinSelectedBeamsToBlockWallsCommand"),

            new ButtonDef(
                "btnAddClunasFromDwg_Utility",
                "Piles From\nDWG",
                "BimTasksV2.Commands.Proxies.AddClunasFromDwgCommand"),

            new ButtonDef(
                "btnFamilyParamChanger",
                "Family Param\nChanger",
                "BimTasksV2.Commands.Proxies.FamilyParamChangerCommand"),

            new ButtonDef(
                "btnTestInfrastructure",
                "Test\nInfrastructure",
                "BimTasksV2.Commands.TestInfrastructureCommand"),
        };
    }
}
