using System.Windows;

namespace PresentationAlive
{
    /// <summary>
    /// Interaction logic for PresentationWindow.xaml
    /// </summary>
    public partial class PresentationWindow : Window
    {
        public PresentationWindow()
        {
            InitializeComponent();
        }

        public void SetContent(UIElement uiElement)
        {
            this.Content = uiElement;
        }
    }
}
