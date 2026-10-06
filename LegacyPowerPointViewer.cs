using System;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickLook.Common.Plugin;

namespace QuickLook.Plugin.OfficeViewer
{
    internal sealed class LegacyPowerPointViewer
    {
        private CancellationTokenSource lifetime;
        private IViewer viewer;
        private PropertyChangedEventHandler titleChanged;
        private ContextObject activeContext;
        public void Prepare(string path, ContextObject context)
        {
            context.SetPreferredSizeFit(new Size(1100, 800), 0.9);
            context.Title = Path.GetFileName(path);
        }
        public async void View(string path, ContextObject context)
        {
            lifetime = new CancellationTokenSource();
            CancellationToken token = lifetime.Token;
            activeContext = context;
            var status = new TextBlock {
                Text = "Preparing legacy PowerPoint preview...\nThe first conversion may take a moment.",
                Margin = new Thickness(30), TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center };
            context.ViewerContent = status;
            context.IsBusy = true;
            Dispatcher ui = status.Dispatcher;
            try
            {
                string preview = await Task.Run(() => LegacyPowerPointConverter.Convert(path, token));
                // Closing or switching documents resets the shared ContextObject.
                token.ThrowIfCancellationRequested();
                await ui.InvokeAsync(() => {
                    if (token.IsCancellationRequested) return;
                    viewer = new QuickLook.Plugin.PDFViewer.Plugin();
                    viewer.Init();
                    // Preserve the original document name in the page counter.
                    titleChanged = (sender, args) => {
                        if (args.PropertyName == "Title" && context.Title != null && context.Title.EndsWith("preview.pdf", StringComparison.Ordinal))
                            context.Title = context.Title.Substring(0, context.Title.Length - "preview.pdf".Length) + Path.GetFileName(path);
                    };
                    context.PropertyChanged += titleChanged;
                    viewer.Prepare(preview, context);
                    viewer.View(preview, context);
                    context.IsBusy = false;
                });
            }
            catch (OperationCanceledException) { }
            catch (Exception error)
            {
                if (token.IsCancellationRequested) return;
                await ui.InvokeAsync(() => {
                    if (token.IsCancellationRequested) return;
                    status.Text = "Could not preview this presentation.\n" + error.Message;
                    context.ViewerContent = status;
                    context.IsBusy = false;
                });
            }
        }
        public void Cleanup()
        {
            if (lifetime != null) { lifetime.Cancel(); lifetime.Dispose(); lifetime = null; }
            if (activeContext != null && titleChanged != null)
                activeContext.PropertyChanged -= titleChanged;
            titleChanged = null;
            activeContext = null;
            if (viewer != null) { viewer.Cleanup(); viewer = null; }
        }
    }
}
