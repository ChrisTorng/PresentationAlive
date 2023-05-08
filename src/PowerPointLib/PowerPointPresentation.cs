using System.Globalization;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Ppt = Microsoft.Office.Interop.PowerPoint;

namespace PresentationAlive.PowerPointLib;

internal sealed class PowerPointPresentation : IDisposable
{
    private readonly Presentation presentation;
    private bool started;
    private bool disposed;

    public const string TAGNAME = "Id";

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

    private void Dispose(bool disposing)
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
        this.started = true;
    }

    internal IEnumerable<PowerPointSubItem> IterateAllSlides()
    {
        foreach (Slide slide in this.presentation.Slides)
        {
            string slideTitle = string.Empty;

            foreach (Ppt.Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        TextRange textRange = shape.TextFrame.TextRange;
                        if (textRange.ParagraphFormat.Bullet.Type == PpBulletType.ppBulletNone &&
                            textRange.Font.Size > 24)
                        {
                            slideTitle = textRange.Text;
                            break;
                        }
                    }
                }
            }

            yield return new PowerPointSubItem(this, slide.SlideIndex, slideTitle);
        }
    }


    internal bool PreviousEnabled =>
        this.started &&
        this.presentation.SlideShowWindow.View.CurrentShowPosition != 1;

    internal bool NextEnabled =>
        this.started &&
        this.presentation.SlideShowWindow.View.CurrentShowPosition !=
            this.presentation.Slides.Count;

    internal bool Previous()
    {
        if (this.PreviousEnabled)
        {
            this.presentation.SlideShowWindow.View.Previous();
            return true;
        }

        return false;
    }

    internal bool Next()
    {
        if (this.NextEnabled)
        {
            this.presentation.SlideShowWindow.View.Next();
            return true;
        }

        return false;
    }

    internal void ShowSlide(int index)
    {
        if (!started)
        {
            this.Start();
        }

        this.presentation.SlideShowWindow.View.GotoSlide(index);
        this.presentation.SlideShowWindow.Activate();
    }

    internal void Stop()
    {
        this.started = false;
        //this.presentation.SlideShowWindow.View.Exit();
    }

    internal void OnSlideShowNextSlide()
    {
        this.SlideShowNextSlide?.Invoke(this, EventArgs.Empty);
    }

    internal void OnSlideShowEnd()
    {
        this.SlideShowEnd?.Invoke(this, EventArgs.Empty);
    }
}
