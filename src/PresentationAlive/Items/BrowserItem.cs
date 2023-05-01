using Microsoft.Web.WebView2.Wpf;
using PresentationAlive.ItemLib;

namespace PresentationAlive.Items
{
    internal class BrowserItem : AbstractItem
    {
        private WebView2? webView2;
        private bool derivedDisposed;

        public BrowserItem(string displayName, string path)
            : base(ItemType.Browser, displayName, path)
        {
        }

        public override void Open()
        {
            this.webView2 = new WebView2
            {
                Source = new Uri(this.Path)
            };
        }

        public override void Start()
        {
            AbstractItem.Start(this.webView2!);
        }

        protected override void Dispose(bool disposing)
        {
            if (!derivedDisposed)
            {
                if (disposing)
                {
                    this.webView2?.Dispose();
                }

                derivedDisposed = true;
            }

            base.Dispose(disposing);
        }
    }
}
