using System;
using System.IO;
using System.Windows;
using PresentationAlive.ItemLib;
using PresentationAlive.PowerPointLib;

namespace PresentationAlive;

public partial class MainWindow : Window
{
    private readonly List<IItem> items;

    public MainWindow()
    {
        this.InitializeComponent();

        this.items = new()
        {
            new PowerPointItem("A", GetFullPath(@"..\data\ppt\a.pptx")),
            new PowerPointItem("B", GetFullPath(@"..\data\ppt\b.pptx")),
        };

        foreach (var item in this.items)
        {
            item.Stopped += this.Item_Stopped;
            this.playList.Items.Add(item.ToString());
        }

        this.playList.SelectedIndex = 0;
    }

    private static string GetFullPath(string file) =>
        Path.Combine(Directory.GetCurrentDirectory(), file);

    private void ButtonStart_Click(object sender, RoutedEventArgs e)
    {
        if (this.playList.SelectedIndex >= 0)
        {
            this.GetItem()?.Start();
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
        Dispatcher.Invoke(() => this.GetItem()?.Close());
    }
}
