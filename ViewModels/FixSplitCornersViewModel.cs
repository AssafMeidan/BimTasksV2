using System;
using System.Windows;
using BimTasksV2.Events;
using BimTasksV2.Helpers.WallSplitter;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    public class FixSplitCornersViewModel : BindableBase
    {
        private readonly ICommandDispatcherService _dispatcher;

        public DelegateCommand FixCornersCommand { get; }

        public FixSplitCornersViewModel()
        {
            var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
            _dispatcher = container.Resolve<ICommandDispatcherService>();

            var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
            eventAgg.GetEvent<BimTasksEvents.FixSplitCornersReadyEvent>()
                .Subscribe(OnSplitComplete, ThreadOption.UIThread);

            FixCornersCommand = new DelegateCommand(ExecuteFixCorners, () => _canFixCorners);
        }

        #region Properties

        private string _statusText = "No split data available. Run Split Wall first.";
        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        private string _buttonText = "Fix Corners";
        public string ButtonText
        {
            get => _buttonText;
            set => SetProperty(ref _buttonText, value);
        }

        private bool _canFixCorners;
        public bool CanFixCorners
        {
            get => _canFixCorners;
            set { SetProperty(ref _canFixCorners, value); FixCornersCommand.RaiseCanExecuteChanged(); }
        }

        private string _resultText = "";
        public string ResultText
        {
            get => _resultText;
            set => SetProperty(ref _resultText, value);
        }

        public Visibility ResultVisibility =>
            string.IsNullOrEmpty(_resultText) ? Visibility.Collapsed : Visibility.Visible;

        #endregion

        private void OnSplitComplete(FixSplitCornersPayload payload)
        {
            StatusText = $"{payload.WallsSplit} wall(s) split into {payload.TotalReplacements} layers.\n\n" +
                         "Dismiss any Revit join errors, then click Fix Corners to trim/extend walls at corner intersections.";
            CanFixCorners = true;
            ResultText = "";
            RaisePropertyChanged(nameof(ResultVisibility));
        }

        private void ExecuteFixCorners()
        {
            CanFixCorners = false;
            ButtonText = "Fixing...";

            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument.Document;
                    string result = WallSplitterEngine.FixCorners(doc);
                    Log.Information("[FixSplitCornersViewModel] {Result}", result);

                    // Update UI on the WPF dispatcher thread
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ResultText = result;
                        RaisePropertyChanged(nameof(ResultVisibility));
                        ButtonText = "Fix Corners";
                        StatusText = "Corner fix complete.";
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[FixSplitCornersViewModel] FixCorners failed");
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ResultText = $"Error: {ex.Message}";
                        RaisePropertyChanged(nameof(ResultVisibility));
                        ButtonText = "Fix Corners";
                        CanFixCorners = true;
                    });
                }
            });
        }
    }
}
