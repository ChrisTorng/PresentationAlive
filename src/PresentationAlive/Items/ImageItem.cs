using System.Windows.Controls;
using System.Windows.Media.Imaging;
using PresentationAlive.ItemLib;

namespace PresentationAlive.Items
{
    internal class ImageItem : IItem
    {
        private static PresentationWindow window = new PresentationWindow();
        private Image? image;
        private bool disposed;

        public event EventHandler? Stopped;

        public ImageItem(string displayName, string path)
        {
            this.DisplayName = displayName;
            this.Path = path;
        }

        ~ImageItem()
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
            "Image: " + this.DisplayName;

        public ItemType ItemType => ItemType.Image;

        public string DisplayName { get; }

        public string Path { get; }

        public bool PreviousEnabled => false;

        public bool NextEnabled => false;


        public void Open()
        {
            this.image = new Image
            {
                Source = new BitmapImage(new Uri(this.Path))
            };
        }

        public void Start()
        {
            window.SetContent(this.image!);
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
        }

        public void Close()
        {
        }
    }
}
