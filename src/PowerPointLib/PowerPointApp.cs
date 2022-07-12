using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;

namespace PresentationAlive.PowerPointLib;

public class PowerPointApp : IDisposable
{
    private static PowerPointApp? instance;
    private readonly PowerPointApplication app;
    private readonly Dictionary<string, PowerPointPresentation> presentations;
    private bool disposed;

    private PowerPointApp()
    {
        app = new()
        {
            Visible = MsoTriState.msoTrue,
            WindowState = PpWindowState.ppWindowMinimized,
        };

        this.ClosePresentationsLefted();

        app.SlideShowNextSlide += this.App_SlideShowNextSlide;
        app.SlideShowEnd += this.App_SlideShowEnd;

        this.presentations = new();
    }

    public static PowerPointApp Instance
    {
        get
        {
            if (instance == null)
            {
                instance = new PowerPointApp();
            }

            return instance;
        }
    }

    public static void DisposeInstance()
    {
        Instance.Dispose();
    }

    ~PowerPointApp()
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

        app.SlideShowNextSlide -= this.App_SlideShowNextSlide;
        app.SlideShowEnd -= this.App_SlideShowEnd;

        if (disposing)
        {
            foreach (var powerPointPresentation in this.presentations.Values)
            {
                powerPointPresentation.Dispose();
            }
        }

        this.ClosePresentationsLefted();

        // app.Quit sometimes has System.Runtime.InteropServices.COMException:
        // 'Application (unknown member) : Invalid request. This operation cannot be performed in this event handler.'
        if (SynchronizationContext.Current != null)
        {
            SynchronizationContext.Current.Post(s => this.app.Quit(), null);
        }
        else
        {
            this.app.Quit();
        }

        // When slide showing, exit app needs GC Collect, or the ppt will not close reliably
        GC.Collect();
        this.disposed = true;
    }

    private void ClosePresentationsLefted()
    {
        // If previous ppt not exited, this can close all lefted
        if (this.app.Presentations.Count > 0)
        {
            foreach (Presentation presentation in this.app.Presentations)
            {
                presentation.Close();
            }
        }
    }

    private static string PresentationId(Presentation presentation) =>
        presentation.Tags[PowerPointPresentation.TAGNAME];

    private void App_SlideShowNextSlide(SlideShowWindow slideShowWindow)
    {
        foreach (var presentation in this.presentations)
        {
            if (presentation.Key ==
                PresentationId(slideShowWindow.Presentation))
            {
                presentation.Value.OnSlideShowNextSlide();
                break;
            }
        }
    }

    private void App_SlideShowEnd(Presentation presentation)
    {
        this.presentations[PresentationId(presentation)].OnSlideShowEnd();
    }

    internal PowerPointPresentation GetPresentation(string path)
    {
        var presentation = app.Presentations.Open(path, WithWindow: MsoTriState.msoFalse);
        PowerPointPresentation powerPointPresentation = new(presentation);
        this.presentations.Add(powerPointPresentation.Id, powerPointPresentation);
        return powerPointPresentation;
    }

    internal void Close(PowerPointPresentation powerPointPresentation)
    {
        foreach (var keyValuePair in this.presentations)
        {
            if (keyValuePair.Value == powerPointPresentation)
            {
                powerPointPresentation.Dispose();
                this.presentations.Remove(keyValuePair.Key);
                break;
            }
        }
    }
}
