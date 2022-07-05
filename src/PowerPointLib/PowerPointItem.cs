using System.Diagnostics.CodeAnalysis;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PresentationAlive.ItemLib;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;

namespace PresentationAlive.PowerPointLib;

public class PowerPointItem : IItem
{
    private static PowerPointApp? app;

    private Presentation? presentation;

    private bool lastSlideReached;

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

    public static void Open()
    {
        if (app == null)
        {
            app = new()
            {
                Visible = MsoTriState.msoTrue,
                WindowState = PpWindowState.ppWindowMinimized,
            };
        }
    }

    public void Start()
    {
        ArgumentNullException.ThrowIfNull(app);

        app.SlideShowEnd += this.App_SlideShowEnd;
        app.SlideShowNextSlide += this.App_SlideShowNextSlide;

        this.presentation = app.Presentations.Open(this.Path, WithWindow: MsoTriState.msoFalse);
        var slideShowSettings = presentation.SlideShowSettings;
        slideShowSettings.ShowWithAnimation = MsoTriState.msoTrue;
        slideShowSettings.Run();
    }

    public void Next()
    {
        ArgumentNullException.ThrowIfNull(app);
        ArgumentNullException.ThrowIfNull(this.presentation);

        if (!this.lastSlideReached)
        {
            this.presentation.SlideShowWindow.View.Next();
        }
        else
        {
            this.presentation.SlideShowWindow.View.Exit();
            //this.Stop();
        }
    }

    private void App_SlideShowNextSlide(SlideShowWindow Wn)
    {
        this.lastSlideReached =
            Wn.View.CurrentShowPosition == this.presentation?.Slides.Count;
    }

    private void App_SlideShowEnd(Presentation Pres)
    {
        this.Stop();
    }

    public void Stop()
    {
        if (app != null)
        {
            app.SlideShowEnd -= this.App_SlideShowEnd;
            app.SlideShowNextSlide -= this.App_SlideShowNextSlide;
        }

        this.presentation?.Close();
        this.presentation = null;

        this.Stopped?.Invoke(this, new EventArgs());
    }

    public static void Close()
    {
        if (app != null)
        {
            app.Quit();
            app = null;
        }
    }
}