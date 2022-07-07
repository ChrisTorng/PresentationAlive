using System.Diagnostics.CodeAnalysis;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PresentationAlive.ItemLib;

namespace PresentationAlive.PowerPointLib;

public class PowerPointItem : IItem
{
    private static readonly PowerPointApp app = PowerPointApp.Instance;
    private PowerPointPresentation? presentation;
    private bool disposed;

    public event EventHandler? Stopped;

    public PowerPointItem(string displayName, string path)
    {
        this.DisplayName = displayName;
        this.Path = path;
    }

    ~PowerPointItem()
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

        if (disposing && this.presentation != null)
        {
            this.presentation.SlideShowNextSlide += this.SlideShowNextSlide;
            this.presentation.SlideShowEnd += this.SlideShowEnd;
            this.presentation.Dispose();
        }

        this.disposed = true;
    }
    private void SlideShowNextSlide(object? sender, EventArgs e)
    {
    }


    private void SlideShowEnd(object? sender, EventArgs e)
    {
        this.Stopped?.Invoke(this, new EventArgs());
    }

    public ItemType ItemType { get; } = ItemType.PowerPoint;

    public string DisplayName { get; }

    public string Path { get; }

    public override string ToString() =>
        "PowerPoint: " + this.DisplayName;

    public void Open()
    {
        this.presentation = app.GetPresentation(this.Path);
        this.presentation.SlideShowNextSlide += this.SlideShowNextSlide;
        this.presentation.SlideShowEnd += this.SlideShowEnd;
    }

    public void Start()
    {
        this.presentation?.Start();
    }

    public void Next()
    {
        this.presentation?.Next();
    }

    public void Stop()
    {
        this.presentation?.Stop();
    }

    public void Close()
    {
        if (this.presentation != null)
        {
            this.presentation.SlideShowNextSlide += this.SlideShowNextSlide;
            this.presentation.SlideShowEnd += this.SlideShowEnd;
            this.presentation.Dispose();
        }
    }
}