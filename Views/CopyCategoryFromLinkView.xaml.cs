using System.Windows;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Dialog for copying elements from a linked Revit model.
    /// Provides cascading selection: Link -> Category -> Family -> Type.
    /// </summary>
    public partial class CopyCategoryFromLinkView : Window
    {
        public CopyCategoryFromLinkView()
        {
            InitializeComponent();
        }
    }
}
