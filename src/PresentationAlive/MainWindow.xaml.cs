using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using PresentationAlive.ItemLib;
using PresentationAlive.PowerPointLib;

namespace PresentationAlive;

public partial class MainWindow : Window
{
    private readonly List<IItem> items;

    public MainWindow()
    {
        this.InitializeComponent();
        this.Closed += MainWindow_Closed;

        PowerPointItem.Open();

        this.items = new()
        {
            new PowerPointItem("A", GetFullPath(@"data\a.pptx")),
            new PowerPointItem("B", GetFullPath(@"data\b.pptx")),
        };

        foreach (var item in this.items)
        {
            item.Stopped += this.Item_Stopped;
            this.playList.Items.Add(item.ToString());
        }

        this.playList.SelectedIndex = 0;
    }

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        foreach (var item in this.items)
        {
            item.Stopped -= this.Item_Stopped;
            item.Stop();
        }

        PowerPointItem.Close();
    }

    private static string GetFullPath(string file) =>
        Path.Combine(Directory.GetCurrentDirectory(), file);

    private void ButtonStart_Click(object sender, RoutedEventArgs e)
    {
        if (this.playList.SelectedIndex >= 0)
        {
            this.GetItem()?.Start();
            this.Activate();
        }
    }

    private IItem? GetItem()
    {
        var item = this.items[this.playList.SelectedIndex];
        return item.ItemType switch
        {
            ItemType.PowerPoint => item as PowerPointItem,
            _ => null
        };
    }

    private void ButtonNext_Click(object sender, RoutedEventArgs e)
    {
        this.GetItem()?.Next();
    }

    private void Item_Stopped(object? sender, EventArgs eventArgs)
    {
        Dispatcher.Invoke(() =>
        {
            if (this.playList.SelectedIndex < this.playList.Items.Count - 1)
            {
                this.playList.SelectedIndex++;
                this.GetItem()?.Start();
                this.Activate();
            }
        });
    }
}
