using System;
using System.Collections.ObjectModel;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    /// <summary>
    /// Info for a single toolbar button.
    /// </summary>
    public class ToolbarButtonInfo
    {
        public string Key { get; set; } = "";
        public string Label { get; set; } = "";
        public string Tooltip { get; set; } = "";
    }

    /// <summary>
    /// A logical group of toolbar buttons (e.g. "Walls", "BOQ").
    /// </summary>
    public class ToolbarGroupInfo
    {
        public string Name { get; set; } = "";
        public ObservableCollection<ToolbarButtonInfo> Buttons { get; set; } = new();
    }

    /// <summary>
    /// ViewModel for the floating toolbar.
    /// Builds button groups and dispatches commands via ICommandDispatcherService.
    /// </summary>
    public class FloatingToolbarViewModel : BindableBase
    {
        private readonly ICommandDispatcherService _dispatcher;

        public ObservableCollection<ToolbarGroupInfo> Groups { get; }
        public DelegateCommand<string> ExecuteCommandCommand { get; }

        public FloatingToolbarViewModel()
        {
            _dispatcher = BimTasksV2.Infrastructure.ContainerLocator.Container
                .Resolve<ICommandDispatcherService>();

            ExecuteCommandCommand = new DelegateCommand<string>(OnExecuteCommand);
            Groups = BuildGroups();
        }

        private void OnExecuteCommand(string? commandKey)
        {
            if (string.IsNullOrEmpty(commandKey)) return;

            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    // Panel-switching commands
                    var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

                    switch (commandKey)
                    {
                        case "FilterTree":
                            eventAgg.GetEvent<Events.BimTasksEvents.SwitchDockablePanelEvent>()
                                .Publish("FilterTree");
                            return;

                        case "CalcAreaVolume":
                            eventAgg.GetEvent<Events.BimTasksEvents.SwitchDockablePanelEvent>()
                                .Publish("ElementCalculation");
                            eventAgg.GetEvent<Events.BimTasksEvents.CalculateElementsEvent>()
                                .Publish(null!);
                            return;
                    }

                    // Try to resolve and execute a command handler by key
                    string handlerTypeName = $"BimTasksV2.Commands.Handlers.{commandKey}Handler";
                    var handlerType = typeof(FloatingToolbarViewModel).Assembly.GetType(handlerTypeName);

                    if (handlerType != null)
                    {
                        var handler = Activator.CreateInstance(handlerType) as Commands.Infrastructure.ICommandHandler;
                        handler?.Execute(uiApp);
                        return;
                    }

                    Log.Warning("[Toolbar] Command not found: {Key}", commandKey);
                    Autodesk.Revit.UI.TaskDialog.Show("BimTasks",
                        $"Command '{commandKey}' is not yet available.");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[Toolbar] Failed to execute {Key}", commandKey);
                    Autodesk.Revit.UI.TaskDialog.Show("Error", $"Command failed:\n{ex.Message}");
                }
            });
        }

        private static ObservableCollection<ToolbarGroupInfo> BuildGroups()
        {
            return new ObservableCollection<ToolbarGroupInfo>
            {
                new ToolbarGroupInfo
                {
                    Name = "Walls",
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        new() { Key = "ChangeWallHeight",                          Label = "Height",     Tooltip = "Change wall height" },
                        new() { Key = "AdjustWallHeightToFloors",                  Label = "To Floors",  Tooltip = "Adjust wall height to floors" },
                        new() { Key = "SetSelectedWindowAndDoorPhaseToWallPhase",  Label = "Sync Phase", Tooltip = "Sync window/door phase to wall" },
                        new() { Key = "AddChipuyToWall",                           Label = "Clad Both",  Tooltip = "Add cladding to both sides" },
                        new() { Key = "AddChipuyToExternalWall",                   Label = "Clad Ext",   Tooltip = "Add external cladding" },
                        new() { Key = "AddChipuyToInternalWall",                   Label = "Clad Int",   Tooltip = "Add internal cladding" },
                        new() { Key = "JoinAllConcreteWallsAndFloorsGeometry",     Label = "Join W+F",   Tooltip = "Join walls and floors geometry" },
                        new() { Key = "JoinSelectedBeamsToBlockWalls",             Label = "Join W+B",   Tooltip = "Join walls and beams geometry" },
                        new() { Key = "CopyLinkedWalls",                           Label = "Copy Link",  Tooltip = "Copy walls from linked model" },
                    }
                },
                new ToolbarGroupInfo
                {
                    Name = "Structure",
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        new() { Key = "SetBeamsToGroundZero",   Label = "Beams Z=0", Tooltip = "Set beams to ground zero" },
                        new() { Key = "AddClunasFromDwg",       Label = "Piles DWG", Tooltip = "Create piles from DWG" },
                    }
                },
                new ToolbarGroupInfo
                {
                    Name = "Schedules",
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        new() { Key = "ExportScheduleToExcel",  Label = "Export",     Tooltip = "Export schedule to Excel" },
                        new() { Key = "EditScheduleInExcel",    Label = "Edit Excel", Tooltip = "Edit schedule in Excel roundtrip" },
                        new() { Key = "CalcAreaVolume",         Label = "Calc A/V",   Tooltip = "Calculate area and volume" },
                    }
                },
                new ToolbarGroupInfo
                {
                    Name = "BOQ",
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        new() { Key = "BOQSetup",            Label = "Setup",      Tooltip = "BOQ setup configuration" },
                        new() { Key = "BOQCreateSchedules",  Label = "Schedules",  Tooltip = "Create BOQ schedules" },
                        new() { Key = "BOQAutoFillQty",      Label = "AutoFill",   Tooltip = "Auto-fill BOQ quantities" },
                        new() { Key = "BOQCalcEffUnitPrice", Label = "Calc Price", Tooltip = "Calculate effective unit price" },
                        new() { Key = "CreateSeifeiChoze",   Label = "Contract",   Tooltip = "Create contract sections" },
                    }
                },
                new ToolbarGroupInfo
                {
                    Name = "Tools",
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        new() { Key = "FilterTree",            Label = "Filter",   Tooltip = "Open filter tree" },
                        new() { Key = "CreateSelectedItem",    Label = "Create",   Tooltip = "Create similar element" },
                        new() { Key = "CreateWindowFamilies",  Label = "Windows",  Tooltip = "Create window families" },
                        new() { Key = "CopyCategoryFromLink",  Label = "Copy Cat", Tooltip = "Copy category from link" },
                    }
                }
            };
        }
    }
}
