using System.Windows;

namespace PresentationAlive
{
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
