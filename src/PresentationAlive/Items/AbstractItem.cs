using System.Windows;
using PresentationAlive.ItemLib;

namespace PresentationAlive.Items
{
    internal abstract class AbstractItem : IItem
    {
        protected static PresentationWindow window = new();
        private bool disposed;

        public event EventHandler? Stopped;

        public AbstractItem(ItemType itemType, string displayName, string path)
        {
            this.ItemType = itemType;
            this.DisplayName = displayName;
            this.Path = path;
        }

        ~AbstractItem()
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

        public override string ToString() =>
            $"{this.ItemType}: {this.DisplayName}";

        public ItemType ItemType { get; }

        public string DisplayName { get; }

        public string Path { get; }

        public bool PreviousEnabled => false;

        public bool NextEnabled => false;


        public abstract void Open();

        public abstract void Start();

        protected static void Start(UIElement uiElement)
        {
            window.SetContent(uiElement);
            window.Show();
            window.Activate();
        }

        public void Previous()
        {
        }

        public void Next()
        {
        }

        public void Stop()
        {
            window.Close();
            this.Stopped?.Invoke(this, new EventArgs());
        }

        public void Close()
        {
        }
    }
}
