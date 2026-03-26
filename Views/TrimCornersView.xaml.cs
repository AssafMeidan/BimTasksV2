using System.Windows;
using System.Windows.Controls;
using BimTasksV2.Services;
using Prism.Ioc;

namespace BimTasksV2.Views
{
    public partial class TrimCornersView : UserControl
    {
        public TrimCornersView()
        {
            InitializeComponent();
        }

        private void TrimCornersButton_Click(object sender, RoutedEventArgs e)
        {
            var dispatcher = Infrastructure.ContainerLocator.Container
                .Resolve<ICommandDispatcherService>();

            dispatcher.Enqueue(uiApp =>
            {
                var handler = new Commands.Handlers.TrimWallCornersHandler();
                handler.Execute(uiApp);
            });
        }
    }
}
