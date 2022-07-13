using System.Windows.Controls;
using System.Windows.Media.Imaging;
using PresentationAlive.ItemLib;

namespace PresentationAlive.Items
{
    internal class ImageItem : AbstractItem
    {
        private Image? image;

        public ImageItem(string displayName, string path)
            : base(ItemType.Image, displayName, path)
        {
        }

        public override void Open()
        {
            this.image = new Image
            {
                Source = new BitmapImage(new Uri(this.Path))
            };
        }

        public override void Start()
        {
            AbstractItem.Start(this.image!);
        }
    }
}
