using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using BimTasksV2.Events;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    /// <summary>
    /// ViewModel for the Element Calculation dashboard.
    /// Displays area and volume statistics for selected Revit elements across 11 categories.
    /// Subscribes to CalculateElementsEvent for triggered recalculation.
    /// </summary>
    public class ElementCalculationViewModel : BindableBase, IDisposable
    {
        #region Fields

        private readonly IRevitContextService _revitContext;
        private readonly IElementCalculationService _calculationService;
        private readonly IEventAggregator _eventAggregator;
        private SubscriptionToken? _calculateEventToken;
        private bool _disposed;

        #endregion

        #region Constructor

        public ElementCalculationViewModel()
        {
            var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
            _revitContext = container.Resolve<IRevitContextService>();
            _calculationService = container.Resolve<IElementCalculationService>();
            _eventAggregator = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

            _calculateEventToken = _eventAggregator
                .GetEvent<BimTasksEvents.CalculateElementsEvent>()
                .Subscribe(_ => ExecuteRecalculate(), ThreadOption.UIThread);

            RecalculateCommand = new DelegateCommand(ExecuteRecalculate);
            CopyToClipboardCommand = new DelegateCommand(ExecuteCopyToClipboard);

            Log.Debug("[ElementCalculationViewModel] Initialized");
        }

        #endregion

        #region Properties - Walls

        private int _wallCount;
        public int WallCount { get => _wallCount; set => SetProperty(ref _wallCount, value); }

        private double _wallArea;
        public double WallArea { get => _wallArea; set => SetProperty(ref _wallArea, value); }

        private double _wallVolume;
        public double WallVolume { get => _wallVolume; set => SetProperty(ref _wallVolume, value); }

        #endregion

        #region Properties - Floors

        private int _floorCount;
        public int FloorCount { get => _floorCount; set => SetProperty(ref _floorCount, value); }

        private double _floorArea;
        public double FloorArea { get => _floorArea; set => SetProperty(ref _floorArea, value); }

        private double _floorVolume;
        public double FloorVolume { get => _floorVolume; set => SetProperty(ref _floorVolume, value); }

        #endregion

        #region Properties - Beams

        private int _beamCount;
        public int BeamCount { get => _beamCount; set => SetProperty(ref _beamCount, value); }

        private double _beamVolume;
        public double BeamVolume { get => _beamVolume; set => SetProperty(ref _beamVolume, value); }

        private double _beamLength;
        public double BeamLength { get => _beamLength; set => SetProperty(ref _beamLength, value); }

        #endregion

        #region Properties - Columns

        private int _columnCount;
        public int ColumnCount { get => _columnCount; set => SetProperty(ref _columnCount, value); }

        private double _columnVolume;
        public double ColumnVolume { get => _columnVolume; set => SetProperty(ref _columnVolume, value); }

        private double _columnHeight;
        public double ColumnHeight { get => _columnHeight; set => SetProperty(ref _columnHeight, value); }

        #endregion

        #region Properties - Foundations

        private int _foundationCount;
        public int FoundationCount { get => _foundationCount; set => SetProperty(ref _foundationCount, value); }

        private double _foundationVolume;
        public double FoundationVolume { get => _foundationVolume; set => SetProperty(ref _foundationVolume, value); }

        private double _foundationArea;
        public double FoundationArea { get => _foundationArea; set => SetProperty(ref _foundationArea, value); }

        #endregion

        #region Properties - Stairs

        private int _stairCount;
        public int StairCount { get => _stairCount; set => SetProperty(ref _stairCount, value); }

        private int _stairRiserCount;
        public int StairRiserCount { get => _stairRiserCount; set => SetProperty(ref _stairRiserCount, value); }

        private double _stairVolume;
        public double StairVolume { get => _stairVolume; set => SetProperty(ref _stairVolume, value); }

        private double _stairUndersideArea;
        public double StairUndersideArea { get => _stairUndersideArea; set => SetProperty(ref _stairUndersideArea, value); }

        private double _stairTreadLength;
        public double StairTreadLength { get => _stairTreadLength; set => SetProperty(ref _stairTreadLength, value); }

        #endregion

        #region Properties - Stair Runs

        private int _stairRunCount;
        public int StairRunCount { get => _stairRunCount; set => SetProperty(ref _stairRunCount, value); }

        private double _stairRunTreadLength;
        public double StairRunTreadLength { get => _stairRunTreadLength; set => SetProperty(ref _stairRunTreadLength, value); }

        #endregion

        #region Properties - Landings

        private int _landingCount;
        public int LandingCount { get => _landingCount; set => SetProperty(ref _landingCount, value); }

        private double _landingArea;
        public double LandingArea { get => _landingArea; set => SetProperty(ref _landingArea, value); }

        private double _landingVolume;
        public double LandingVolume { get => _landingVolume; set => SetProperty(ref _landingVolume, value); }

        #endregion

        #region Properties - Railings

        private int _railingCount;
        public int RailingCount { get => _railingCount; set => SetProperty(ref _railingCount, value); }

        private double _railingLength;
        public double RailingLength { get => _railingLength; set => SetProperty(ref _railingLength, value); }

        #endregion

        #region Properties - Doors

        private int _doorCount;
        public int DoorCount { get => _doorCount; set => SetProperty(ref _doorCount, value); }

        #endregion

        #region Properties - Windows

        private int _windowCount;
        public int WindowCount { get => _windowCount; set => SetProperty(ref _windowCount, value); }

        #endregion

        #region Properties - Totals

        private double _totalConcreteVolume;
        public double TotalConcreteVolume { get => _totalConcreteVolume; set => SetProperty(ref _totalConcreteVolume, value); }

        private int _totalElementCount;
        public int TotalElementCount { get => _totalElementCount; set => SetProperty(ref _totalElementCount, value); }

        private int _selectedElementCount;
        public int SelectedElementCount { get => _selectedElementCount; set => SetProperty(ref _selectedElementCount, value); }

        #endregion

        #region Properties - State

        private bool _isCalculating;
        public bool IsCalculating { get => _isCalculating; set => SetProperty(ref _isCalculating, value); }

        private bool _hasData;
        public bool HasData { get => _hasData; set => SetProperty(ref _hasData, value); }

        private string _statusMessage = "";
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        private long _calculationTimeMs;
        public long CalculationTimeMs { get => _calculationTimeMs; set => SetProperty(ref _calculationTimeMs, value); }

        #endregion

        #region Commands

        public DelegateCommand RecalculateCommand { get; }
        public DelegateCommand CopyToClipboardCommand { get; }

        #endregion

        #region Methods

        private void ExecuteRecalculate()
        {
            Log.Debug("[ElementCalculationViewModel] Recalculate requested");

            IsCalculating = true;
            StatusMessage = "Calculating...";
            ResetValues();

            try
            {
                var uiDoc = _revitContext.UIDocument;
                if (uiDoc == null)
                {
                    StatusMessage = "Error: No active Revit document";
                    return;
                }

                var doc = uiDoc.Document;
                ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
                var result = _calculationService.CalculateElements(doc, selectedIds);

                // Map result to properties
                WallCount = result.WallCount;
                WallArea = result.WallAreaM2;
                WallVolume = result.WallVolumeM3;

                FloorCount = result.FloorCount;
                FloorArea = result.FloorAreaM2;
                FloorVolume = result.FloorVolumeM3;

                BeamCount = result.BeamCount;
                BeamVolume = result.BeamVolumeM3;
                BeamLength = result.BeamLengthM;

                ColumnCount = result.ColumnCount;
                ColumnVolume = result.ColumnVolumeM3;
                ColumnHeight = result.ColumnHeightM;

                FoundationCount = result.FoundationCount;
                FoundationVolume = result.FoundationVolumeM3;
                FoundationArea = result.FoundationAreaM2;

                StairCount = result.StairCount;
                StairRiserCount = result.StairRiserCount;
                StairVolume = result.StairVolumeM3;
                StairUndersideArea = result.StairUndersideAreaM2;
                StairTreadLength = result.StairTreadLengthM;

                StairRunCount = result.StairRunCount;
                StairRunTreadLength = result.StairRunTreadLengthM;

                LandingCount = result.LandingCount;
                LandingArea = result.LandingAreaM2;
                LandingVolume = result.LandingVolumeM3;

                RailingCount = result.RailingCount;
                RailingLength = result.RailingLengthM;

                DoorCount = result.DoorCount;
                WindowCount = result.WindowCount;

                TotalConcreteVolume = result.TotalConcreteVolumeM3;
                TotalElementCount = result.TotalElementCount;
                SelectedElementCount = result.SelectedElementCount;
                CalculationTimeMs = result.CalculationTimeMs;

                HasData = result.TotalElementCount > 0;
                StatusMessage = result.HasErrors
                    ? $"Completed with errors ({result.Errors.Count})"
                    : $"Calculated in {result.CalculationTimeMs}ms";

                Log.Information("[ElementCalculationViewModel] Complete. Total: {Volume:F2} m3", TotalConcreteVolume);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[ElementCalculationViewModel] Calculation failed");
                StatusMessage = $"Error: {ex.Message}";
                HasData = false;
            }
            finally
            {
                IsCalculating = false;
            }
        }

        private void ExecuteCopyToClipboard()
        {
            try
            {
                var text = FormatResultsForClipboard();
                System.Windows.Clipboard.SetText(text);
                StatusMessage = "Copied to clipboard";
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[ElementCalculationViewModel] Copy to clipboard failed");
                StatusMessage = "Copy failed";
            }
        }

        private string FormatResultsForClipboard()
        {
            return $@"=== Element Calculation ===

Walls: {WallCount}
  Area: {WallArea:F2} m2
  Volume: {WallVolume:F2} m3

Floors: {FloorCount}
  Area: {FloorArea:F2} m2
  Volume: {FloorVolume:F2} m3

Beams: {BeamCount}
  Volume: {BeamVolume:F2} m3
  Length: {BeamLength:F2} m

Columns: {ColumnCount}
  Volume: {ColumnVolume:F2} m3
  Height: {ColumnHeight:F2} m

Foundations: {FoundationCount}
  Area: {FoundationArea:F2} m2
  Volume: {FoundationVolume:F2} m3

Stairs: {StairCount} ({StairRiserCount} risers)
  Concrete Volume: {StairVolume:F2} m3
  Tread Length: {StairTreadLength:F2} m
  Underside Area: {StairUndersideArea:F2} m2

Stair Runs: {StairRunCount}
  Tread Length: {StairRunTreadLength:F2} m

Landings: {LandingCount}
  Flooring Area: {LandingArea:F2} m2
  Volume: {LandingVolume:F2} m3

Railings: {RailingCount}
  Length: {RailingLength:F2} m

Doors: {DoorCount}
Windows: {WindowCount}

========================================
Total Concrete: {TotalConcreteVolume:F2} m3
========================================";
        }

        private void ResetValues()
        {
            WallCount = 0; WallArea = 0; WallVolume = 0;
            FloorCount = 0; FloorArea = 0; FloorVolume = 0;
            BeamCount = 0; BeamVolume = 0; BeamLength = 0;
            ColumnCount = 0; ColumnVolume = 0; ColumnHeight = 0;
            FoundationCount = 0; FoundationVolume = 0; FoundationArea = 0;
            StairCount = 0; StairRiserCount = 0; StairVolume = 0;
            StairUndersideArea = 0; StairTreadLength = 0;
            StairRunCount = 0; StairRunTreadLength = 0;
            LandingCount = 0; LandingArea = 0; LandingVolume = 0;
            RailingCount = 0; RailingLength = 0;
            DoorCount = 0; WindowCount = 0;
            TotalConcreteVolume = 0; TotalElementCount = 0;
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                if (_calculateEventToken != null)
                {
                    _eventAggregator.GetEvent<BimTasksEvents.CalculateElementsEvent>()
                        .Unsubscribe(_calculateEventToken);
                    _calculateEventToken = null;
                }
            }

            _disposed = true;
        }

        #endregion
    }
}
