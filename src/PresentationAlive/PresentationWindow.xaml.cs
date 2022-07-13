using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

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
