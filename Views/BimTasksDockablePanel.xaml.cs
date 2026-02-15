using System.Windows.Controls;
using Autodesk.Revit.UI;

namespace BimTasksV2.Views
{
    /// <summary>
    /// A dockable Page that hosts switchable content in the Revit UI.
    /// Implements IDockablePaneProvider so Revit can register it as a dockable pane.
    /// </summary>
    public partial class BimTasksDockablePanel : Page, IDockablePaneProvider
    {
        public BimTasksDockablePanel()
        {
            InitializeComponent();
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Right
            };
        }
    }
}
