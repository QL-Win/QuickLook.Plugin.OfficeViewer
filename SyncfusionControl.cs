using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
using Syncfusion.Windows.Controls.RichTextBoxAdv;

namespace QuickLook.Plugin.OfficeViewer
{
    static class SyncfusionControl
    {
        public static Control OpenWord(string path)
        {
            var editor = new SfRichTextBoxAdv
            {
                IsReadOnly = true, Background = Brushes.Transparent, EnableMiniToolBar = false
            };

            editor.LoadAsync(path);

            return editor;
        }
    }
}
