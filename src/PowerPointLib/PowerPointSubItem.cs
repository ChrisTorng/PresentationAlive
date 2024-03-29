﻿using System.Globalization;
using PresentationAlive.ItemLib;

namespace PresentationAlive.PowerPointLib
{
    public class PowerPointSubItem : IItem
    {
        private readonly PowerPointPresentation presentation;
        private readonly int index;
        private bool disposed;

        internal PowerPointSubItem(PowerPointPresentation presentation, int index, string displayName)
        {
            this.presentation = presentation;
            this.index = index;
            this.Path = index.ToString(CultureInfo.InvariantCulture);
            this.DisplayName = displayName;
        }

        public ItemType ItemType => ItemType.PowerPointSlide;

        public string DisplayName { get; }

        public string Path { get; }

        public IEnumerable<IItem>? SubItems { get; }

        public override string ToString() =>
            $"    {this.Path}: {this.DisplayName}";

#pragma warning disable CS0067 // The event 'PowerPointSubItem.Stopped' is never used
        public event EventHandler? Stopped;
#pragma warning restore CS0067 // The event 'PowerPointSubItem.Stopped' is never used

        public void Close()
        {
        }

        ~PowerPointSubItem()
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

            this.disposed = true;
        }

        public void Start()
        {
            this.presentation.ShowSlide(this.index);
        }

        public void Open()
        {
        }

        public void Stop()
        {
        }
    }
}
