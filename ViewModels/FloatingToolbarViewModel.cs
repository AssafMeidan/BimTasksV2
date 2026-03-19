using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using BimTasksV2.Events;
using BimTasksV2.Ribbon;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    public class ToolbarButtonInfo
    {
        public string Key { get; set; } = "";
        public string Label { get; set; } = "";
        public string Tooltip { get; set; } = "";
        public BitmapSource? Icon { get; set; }
    }

    public class ToolbarGroupInfo : BindableBase
    {
        public string Name { get; set; } = "";
        public SolidColorBrush AccentBrush { get; set; } = Brushes.Gray;
        public ObservableCollection<ToolbarButtonInfo> Buttons { get; set; } = new();

        private bool _isExpanded = true;
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (SetProperty(ref _isExpanded, value))
                    RaisePropertyChanged(nameof(ExpanderIcon));
            }
        }

        public string ExpanderIcon => _isExpanded ? "\u25BE" : "\u25B8";
    }

    public class FloatingToolbarViewModel : BindableBase
    {
        private readonly ICommandDispatcherService _dispatcher;

        private bool _isVertical;
        public bool IsVertical
        {
            get => _isVertical;
            set
            {
                if (SetProperty(ref _isVertical, value))
                {
                    RaisePropertyChanged(nameof(GroupsOrientation));
                    RaisePropertyChanged(nameof(ButtonsOrientation));
                    RaisePropertyChanged(nameof(OrientationIcon));
                    RaisePropertyChanged(nameof(HScrollVisibility));
                    RaisePropertyChanged(nameof(VScrollVisibility));
                }
            }
        }

        public Orientation GroupsOrientation => _isVertical ? Orientation.Vertical : Orientation.Horizontal;
        public Orientation ButtonsOrientation => _isVertical ? Orientation.Vertical : Orientation.Horizontal;
        public string OrientationIcon => _isVertical ? "\u2194" : "\u2195";
        public ScrollBarVisibility HScrollVisibility => _isVertical ? ScrollBarVisibility.Disabled : ScrollBarVisibility.Auto;
        public ScrollBarVisibility VScrollVisibility => _isVertical ? ScrollBarVisibility.Auto : ScrollBarVisibility.Disabled;

        public ObservableCollection<ToolbarGroupInfo> Groups { get; }
        public DelegateCommand<string> ExecuteCommandCommand { get; }
        public DelegateCommand ToggleOrientationCommand { get; }
        public DelegateCommand<ToolbarGroupInfo> ToggleGroupCommand { get; }

        /// <summary>
        /// Raised when orientation changes so the Window can adjust layout.
        /// </summary>
        public event Action? OrientationChanged;

        // Command key -> IconGenerator ribbon button name
        private static readonly Dictionary<string, string> CommandToIconKey = new()
        {
            ["ChangeWallHeight"] = "btnChangeWallHeight",
            ["AdjustWallHeightToFloors"] = "btnAdjustWallHeight",
            ["SetSelectedWindowAndDoorPhaseToWallPhase"] = "btnSetWindowDoorPhase",
            ["AddChipuyToWall"] = "btnAddChipuyToWall",
            ["AddChipuyToExternalWall"] = "btnAddChipuyExternal",
            ["AddChipuyToInternalWall"] = "btnAddChipuyInternal",
            ["JoinAllConcreteWallsAndFloorsGeometry"] = "btnJoinConcreteWallsFloors",
            ["JoinSelectedBeamsToBlockWalls"] = "btnJoinBeamsToBlockWalls",
            ["CopyLinkedWalls"] = "btnCopyLinkedWalls",
            ["SetBeamsToGroundZero"] = "btnSetBeamsToGround",
            ["AddClunasFromDwg"] = "btnAddClunasFromDwg",
            ["ExportScheduleToExcel"] = "btnExportScheduleToExcel",
            ["EditScheduleInExcel"] = "btnEditScheduleInExcel",
            ["CalcAreaVolume"] = "btnCheckWallsAreaVolume",
            ["BOQSetup"] = "btnBOQSetup",
            ["BOQCreateSchedules"] = "btnBOQCreateSchedules",
            ["BOQAutoFillQty"] = "btnBOQAutoFillQty",
            ["BOQCalcEffUnitPrice"] = "btnBOQCalcEffUnitPrice",
            ["CreateSeifeiChoze"] = "btnCreateSeifeiChoze",
            ["FilterTree"] = "btnShowFilterTree",
            ["CreateSelectedItem"] = "btnCreateSelectedItem",
            ["CreateWindowFamilies"] = "btnCreateWindowFamilies",
            ["CopyCategoryFromLink"] = "btnCopyCategoryFromLink",
        };

        // Settings persistence
        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BimTasksV2");
        private static readonly string SettingsPath = Path.Combine(SettingsDir, "toolbar-settings.json");

        public FloatingToolbarViewModel()
        {
            _dispatcher = Infrastructure.ContainerLocator.Container
                .Resolve<ICommandDispatcherService>();

            ExecuteCommandCommand = new DelegateCommand<string>(OnExecuteCommand);
            ToggleOrientationCommand = new DelegateCommand(() =>
            {
                IsVertical = !IsVertical;
                OrientationChanged?.Invoke();
            });
            ToggleGroupCommand = new DelegateCommand<ToolbarGroupInfo>(g =>
            {
                if (g != null) g.IsExpanded = !g.IsExpanded;
            });

            Groups = BuildGroups();
        }

        #region Command Execution

        private void OnExecuteCommand(string? commandKey)
        {
            if (string.IsNullOrEmpty(commandKey)) return;

            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    // Panel-switching commands
                    var eventAgg = Infrastructure.ContainerLocator.EventAggregator;

                    switch (commandKey)
                    {
                        case "FilterTree":
                            eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                                .Publish("FilterTree");
                            return;

                        case "CalcAreaVolume":
                            eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                                .Publish("ElementCalculation");
                            eventAgg.GetEvent<BimTasksEvents.CalculateElementsEvent>()
                                .Publish(null!);
                            return;
                    }

                    // Resolve and execute a command handler by key
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

        #endregion

        #region Settings Persistence

        public void SaveSettings(double left, double top)
        {
            try
            {
                var settings = new ToolbarSettings
                {
                    Left = left,
                    Top = top,
                    IsVertical = _isVertical,
                    CollapsedGroups = Groups.Where(g => !g.IsExpanded).Select(g => g.Name).ToArray()
                };

                Directory.CreateDirectory(SettingsDir);
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath, json);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[Toolbar] Failed to save settings");
            }
        }

        public (double? left, double? top) LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath))
                    return (null, null);

                var json = File.ReadAllText(SettingsPath);
                var settings = JsonSerializer.Deserialize<ToolbarSettings>(json);
                if (settings == null)
                    return (null, null);

                IsVertical = settings.IsVertical;

                var collapsed = new HashSet<string>(settings.CollapsedGroups ?? Array.Empty<string>());
                foreach (var group in Groups)
                    group.IsExpanded = !collapsed.Contains(group.Name);

                return (settings.Left, settings.Top);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[Toolbar] Failed to load settings");
                return (null, null);
            }
        }

        private class ToolbarSettings
        {
            public double? Left { get; set; }
            public double? Top { get; set; }
            public bool IsVertical { get; set; }
            public string[] CollapsedGroups { get; set; } = Array.Empty<string>();
        }

        #endregion

        #region Group Building

        private static SolidColorBrush FreezeBrush(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        // Same colors as IconGenerator panel backgrounds
        private static readonly SolidColorBrush WallAccent = FreezeBrush(59, 130, 246);
        private static readonly SolidColorBrush StructAccent = FreezeBrush(245, 158, 11);
        private static readonly SolidColorBrush SchedAccent = FreezeBrush(16, 185, 129);
        private static readonly SolidColorBrush BoqAccent = FreezeBrush(139, 92, 246);
        private static readonly SolidColorBrush ToolAccent = FreezeBrush(6, 182, 212);

        private static BitmapSource? GetButtonIcon(string commandKey)
        {
            if (CommandToIconKey.TryGetValue(commandKey, out var iconKey))
                return IconGenerator.GetIcon(iconKey, 20);
            return null;
        }

        private static ToolbarButtonInfo Btn(string key, string label, string tooltip)
        {
            return new ToolbarButtonInfo
            {
                Key = key,
                Label = label,
                Tooltip = tooltip,
                Icon = GetButtonIcon(key)
            };
        }

        private static ObservableCollection<ToolbarGroupInfo> BuildGroups()
        {
            return new ObservableCollection<ToolbarGroupInfo>
            {
                new ToolbarGroupInfo
                {
                    Name = "Walls",
                    AccentBrush = WallAccent,
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        Btn("ChangeWallHeight",                          "Height",     "Change wall height"),
                        Btn("AdjustWallHeightToFloors",                  "To Floors",  "Adjust wall height to floors"),
                        Btn("SetSelectedWindowAndDoorPhaseToWallPhase",  "Sync Phase", "Sync window/door phase to wall"),
                        Btn("AddChipuyToWall",                           "Clad Both",  "Add cladding to both sides"),
                        Btn("AddChipuyToExternalWall",                   "Clad Ext",   "Add external cladding"),
                        Btn("AddChipuyToInternalWall",                   "Clad Int",   "Add internal cladding"),
                        Btn("JoinAllConcreteWallsAndFloorsGeometry",     "Join W+F",   "Join walls and floors geometry"),
                        Btn("JoinSelectedBeamsToBlockWalls",             "Join W+B",   "Join walls and beams geometry"),
                        Btn("CopyLinkedWalls",                           "Copy Link",  "Copy walls from linked model"),
                    }
                },
                new ToolbarGroupInfo
                {
                    Name = "Structure",
                    AccentBrush = StructAccent,
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        Btn("SetBeamsToGroundZero",   "Beams Z=0", "Set beams to ground zero"),
                        Btn("AddClunasFromDwg",       "Piles DWG", "Create piles from DWG"),
                    }
                },
                new ToolbarGroupInfo
                {
                    Name = "Schedules",
                    AccentBrush = SchedAccent,
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        Btn("ExportScheduleToExcel",  "Export",     "Export schedule to Excel"),
                        Btn("EditScheduleInExcel",    "Edit Excel", "Edit schedule in Excel roundtrip"),
                        Btn("CalcAreaVolume",         "Calc A/V",   "Calculate area and volume"),
                    }
                },
                new ToolbarGroupInfo
                {
                    Name = "BOQ",
                    AccentBrush = BoqAccent,
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        Btn("BOQSetup",            "Setup",      "BOQ setup configuration"),
                        Btn("BOQCreateSchedules",  "Schedules",  "Create BOQ schedules"),
                        Btn("BOQAutoFillQty",      "AutoFill",   "Auto-fill BOQ quantities"),
                        Btn("BOQCalcEffUnitPrice", "Calc Price", "Calculate effective unit price"),
                        Btn("CreateSeifeiChoze",   "Contract",   "Create contract sections"),
                    }
                },
                new ToolbarGroupInfo
                {
                    Name = "Tools",
                    AccentBrush = ToolAccent,
                    Buttons = new ObservableCollection<ToolbarButtonInfo>
                    {
                        Btn("FilterTree",            "Filter",   "Open filter tree"),
                        Btn("CreateSelectedItem",    "Create",   "Create similar element"),
                        Btn("CreateWindowFamilies",  "Windows",  "Create window families"),
                        Btn("CopyCategoryFromLink",  "Copy Cat", "Copy category from link"),
                    }
                }
            };
        }

        #endregion
    }
}
