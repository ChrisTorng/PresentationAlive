using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PresentationAlive.PowerPointLib;

internal class PowerPointPresentation : IDisposable
{
    private readonly Presentation presentation;
    private bool lastSlideReached;
    private bool disposed;

    public static readonly string TAGNAME = "Id";

    public string Id { get; }
    public event EventHandler? SlideShowNextSlide;
    public event EventHandler? SlideShowEnd;

    public PowerPointPresentation(Presentation presentation)
    {
        this.Id = presentation.FullName;
        presentation.Tags.Add(TAGNAME, this.Id);
        this.presentation = presentation;
    }

    ~PowerPointPresentation()
    {
        this.Dispose(false);
    }

    public void Dispose()
    {
        this.Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (this.disposed)
        {
            return;
        }

        this.presentation.Close();
        this.disposed = true;
    }

    internal void Start()
    {
        var slideShowSettings = presentation.SlideShowSettings;
        slideShowSettings.ShowWithAnimation = MsoTriState.msoTrue;
        slideShowSettings.Run();
    }

    internal bool Next()
    {
        if (!this.lastSlideReached)
        {
            this.presentation.SlideShowWindow.View.Next();
        }

        return !this.lastSlideReached;
    }

    internal void Stop()
    {
        //this.presentation.SlideShowWindow.View.Exit();
    }

    internal void OnSlideShowNextSlide()
    {
        this.lastSlideReached =
            this.presentation.SlideShowWindow.View.CurrentShowPosition ==
            this.presentation.Slides.Count;

        this.SlideShowNextSlide?.Invoke(this, EventArgs.Empty);
    }

    internal void OnSlideShowEnd()
    {
        this.SlideShowEnd?.Invoke(this, EventArgs.Empty);
    }
}
