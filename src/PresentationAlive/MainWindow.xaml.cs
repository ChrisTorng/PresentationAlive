using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Threading;
using PresentationAlive.ItemLib;
using PresentationAlive.Items;
using PresentationAlive.PowerPointLib;

namespace PresentationAlive;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;
    private readonly List<IItem> items;

    public MainWindow()
    {
        this.InitializeComponent();
        this.DataContext = this;
        this.Closed += MainWindow_Closed;

        //_ = PowerPointApp.Instance;

        this.items = new()
        {
            new ImageItem("Image1", GetFullPath(@"data\Image1.png")),
            new ImageItem("Image2", GetFullPath(@"data\Image2.jpg")),
            new BrowserItem("就是這個時刻", "https://www.youtube.com/watch?v=8xGdaxTpAYA"),
            //new PowerPointItem("A", GetFullPath(@"data\a.pptx")),
            //new PowerPointItem("B", GetFullPath(@"data\b.pptx")),
        };

        foreach (var item in this.items)
        {
            item.Stopped += this.Item_Stopped;
            item.Open();
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

        //PowerPointApp.DisposeInstance();
    }

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    private void BindingChanged()
    {
        this.OnPropertyChanged(nameof(this.PreviousEnabled));
        this.OnPropertyChanged(nameof(this.NextEnabled));
    }

    private static string GetFullPath(string file) =>
        Path.Combine(Directory.GetCurrentDirectory(), file);

    private void ButtonStart_Click(object sender, RoutedEventArgs e)
    {
        if (this.playList.SelectedIndex >= 0)
        {
            this.GetItem()?.Start();
            this.Activate();
            this.BindingChanged();
        }
    }

    private IItem? GetItem()
    {
        var item = this.items[this.playList.SelectedIndex];
        return item.ItemType switch
        {
            ItemType.PowerPoint => item as PowerPointItem,
            ItemType.Image => item as ImageItem,
            ItemType.Browser => item as BrowserItem,
            _ => null
        };
    }

    public bool PreviousEnabled =>
        this.CurrentItemPreviousAvailable || this.PreviousItemAvailable;

    private bool CurrentItemPreviousAvailable =>
        (this.GetItem()?.PreviousEnabled).GetValueOrDefault();

    private bool PreviousItemAvailable =>
        this.playList.SelectedIndex > 0;

    public bool NextEnabled =>
        this.CurrentItemNextAvailable || this.NextItemAvailable;

    private bool CurrentItemNextAvailable =>
        (this.GetItem()?.NextEnabled).GetValueOrDefault();

    private bool NextItemAvailable =>
        this.playList.SelectedIndex < this.playList.Items.Count - 1;

    private void ButtonPrevious_Click(object sender, RoutedEventArgs e)
    {
        if (this.CurrentItemPreviousAvailable)
        {
            this.GetItem()?.Previous();
        }
        else
        {
            this.playList.SelectedIndex--;
            this.GetItem()?.Start();
            this.Activate();
        }

        this.BindingChanged();
    }

    private void ButtonNext_Click(object sender, RoutedEventArgs e)
    {
        if (this.CurrentItemNextAvailable)
        {
            this.GetItem()?.Next();
        }
        else
        {
            this.playList.SelectedIndex++;
            this.GetItem()?.Start();
            this.Activate();
        }

        this.BindingChanged();
    }

    private void ButtonStop_Click(object sender, RoutedEventArgs e)
    {
        this.BindingChanged();
    }

    private void Item_Stopped(object? sender, EventArgs eventArgs)
    {
        Dispatcher.Invoke(() =>
        {
            if (this.NextItemAvailable)
            {
                this.playList.SelectedIndex++;
                this.GetItem()?.Start();
                this.Activate();
                this.BindingChanged();
            }
        });
    }
}
