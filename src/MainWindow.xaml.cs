using System;
using System.IO;
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
        PowerPointApp app;
        Presentation? presentation;

        public MainWindow()
        {
            InitializeComponent();
            this.Closing += MainWindow_Closing;

            this.app = new()
            {
               Visible = MsoTriState.msoTrue,
               WindowState = PpWindowState.ppWindowMinimized,
            };
            app.SlideShowEnd += this.App_SlideShowEnd;

            this.playList.Items.Add(GetFullPath(@"..\data\ppt\a.pptx"));
            this.playList.Items.Add(GetFullPath(@"..\data\ppt\b.pptx"));
            this.playList.SelectedIndex = 0;
        }

        private static string GetFullPath(string file) =>
            Path.Combine(Directory.GetCurrentDirectory(), file);

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            this.presentation?.Close();
            this.presentation = null;
            this.app?.Quit();
            this.app = null!;
            GC.Collect();
        }

        private void ButtonStart_Click(object sender, RoutedEventArgs e)
        {
            this.StartSlideShow();
        }

        private void StartSlideShow()
        {
            if (this.playList.SelectedItem != null)
            {
                this.presentation = app.Presentations.Open(this.playList.SelectedItem.ToString());
                var slideShowSettings = presentation.SlideShowSettings;
                slideShowSettings.Run();
            }
        }

        private void App_SlideShowEnd(Presentation Pres)
        {
            this.presentation?.Close();
            this.presentation = null;

            Dispatcher.Invoke(() => {
                if (this.playList.SelectedIndex < this.playList.Items.Count - 1)
                {
                    this.playList.SelectedIndex++;
                    this.StartSlideShow();
                }
            });
        }

        private void ButtonNext_Click(object sender, RoutedEventArgs e)
        {
            this.presentation?.SlideShowWindow.View.Next();
        }
    }
}
