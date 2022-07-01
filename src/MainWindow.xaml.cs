using System;
using System.IO;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;

namespace PresentationAlive;

public partial class MainWindow : Window
{
    List<IItem> items;

    PowerPointApp app;
    Presentation? presentation;

    public MainWindow()
    {
        this.InitializeComponent();
        this.Closing += MainWindow_Closing;

        this.app = new()
        {
            Visible = MsoTriState.msoTrue,
            WindowState = PpWindowState.ppWindowMinimized,
        };
        app.SlideShowEnd += this.App_SlideShowEnd;

        this.items = new()
        {
            new PowerPointItem("A", GetFullPath(@"..\data\ppt\a.pptx")),
            new PowerPointItem("B", GetFullPath(@"..\data\ppt\b.pptx")),
        };

        foreach (var item in this.items)
        {
            this.playList.Items.Add(item.ToString());
        }

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
        this.StartSlideShow(this.items[0] as PowerPointItem);
    }

    private void StartSlideShow(PowerPointItem item)
    {
        if (this.playList.SelectedItem != null)
        {
            this.presentation = app.Presentations.Open(item.Path);
            var slideShowSettings = presentation.SlideShowSettings;
            slideShowSettings.Run();
        }
    }

    private void App_SlideShowEnd(Presentation Pres)
    {
        this.presentation?.Close();
        this.presentation = null;

        Dispatcher.Invoke(() =>
        {
            if (this.playList.SelectedIndex < this.items.Count - 1)
            {
                this.playList.SelectedIndex++;
                this.StartSlideShow(this.items[this.playList.SelectedIndex] as PowerPointItem);
            }
        });
    }

    private void ButtonNext_Click(object sender, RoutedEventArgs e)
    {
        this.presentation?.SlideShowWindow.View.Next();
    }
}
