using System;
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
        PowerPointApp? app = new()
        {
            WindowState = PpWindowState.ppWindowMinimized,
            Visible = MsoTriState.msoTrue
        };
        Presentation? presentation;

        public MainWindow()
        {
            InitializeComponent();
            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            this.app?.Quit();
            this.app = null;
            GC.Collect();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.presentation = app.Presentations.Open("D:\\Users\\ChrisTorng\\Documents\\個人\\教會\\UP\\20220213\\恢復起初的愛-欣仁版本 - Copy.pptx");
            var slideShowSettings = presentation.SlideShowSettings;
            slideShowSettings.Run();
            app.SlideShowEnd += App_SlideShowEnd1;
        }

        private void App_SlideShowEnd1(Presentation Pres)
        {
            app.SlideShowEnd -= App_SlideShowEnd1;
            this.presentation?.Close();
            this.presentation = app.Presentations.Open("D:\\Users\\ChrisTorng\\Documents\\個人\\教會\\UP\\20220213\\恢復起初的愛 - Copy.pptx");
            var slideShowSettings = presentation.SlideShowSettings;
            slideShowSettings.Run();
            app.SlideShowEnd += App_SlideShowEnd2;
        }

        private void App_SlideShowEnd2(Presentation Pres)
        {
            app.SlideShowEnd -= App_SlideShowEnd2;
            this.presentation?.Close();
            this.presentation = null;
        }
    }
}
