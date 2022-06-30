using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;

namespace src
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var app = new PowerPointApp()
            {
                WindowState = PpWindowState.ppWindowMinimized,
                Visible = MsoTriState.msoTrue
            };
            var presentation = app.Presentations.Open("D:\\Users\\ChrisTorng\\Documents\\個人\\教會\\UP\\20220213\\恢復起初的愛-欣仁版本 - Copy.pptx");
            var slideShowSettings = presentation.SlideShowSettings;
            slideShowSettings.Run();

        }
    }
}
