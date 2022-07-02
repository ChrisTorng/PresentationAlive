using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;

namespace PresentationAlive.PowerPointLib;

public class PowerPointItem : IItem
{
    private PowerPointApp? app;
    private Presentation? presentation;

    public event EventHandler? Stopped;

    public PowerPointItem(string displayName, string path)
    {
        this.DisplayName = displayName;
        this.Path = path;
    }

    public ItemType ItemType { get; } = ItemType.PowerPoint;

    public string DisplayName { get; }

    public string Path { get; }

    public override string ToString() =>
        "PowerPoint: " + this.DisplayName;

    public void Start()
    {
        this.app = new()
        {
            Visible = MsoTriState.msoTrue,
            WindowState = PpWindowState.ppWindowMinimized,
        };
        app.SlideShowEnd += this.App_SlideShowEnd;

        this.presentation = app.Presentations.Open(this.Path);
        var slideShowSettings = presentation.SlideShowSettings;
        slideShowSettings.Run();
    }

    public void Next()
    {
        this.presentation?.SlideShowWindow.View.Next();
    }

    private void App_SlideShowEnd(Presentation Pres)
    {
        if (this.Stopped != null)
        {
            this.Stopped(this, new EventArgs());
        }
    }

    public void Close()
    {
        this.presentation?.Close();
        this.presentation = null;
        this.app?.Quit();
        this.app = null!;
    }
}