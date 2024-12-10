using QuickLook.Plugin.PDFViewer;
using Syncfusion;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Spreadsheet;
using Syncfusion.UI.Xaml.Spreadsheet.GraphicCells;
using Syncfusion.UI.Xaml.SpreadsheetHelper;
using Syncfusion.Windows.Controls.RichTextBoxAdv;
using System;
using System.IO;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace QuickLook.Plugin.OfficeViewer;

internal static class SyncfusionControl
{
    public static Control Open(string path)
    {
        if (new FileInfo(path).IsReadOnly)
        {
            return new Label { Content = "Readonly file is not supported." };
        }

        return (Path.GetExtension(path)?.ToLower()) switch
        {
            ".doc" or ".docx" or ".docm" or ".rtf" => OpenWord(path),
            ".xls" or ".xlsx" or ".xlsm" => OpenExcel(path),
            ".pptx" or ".pptm" or ".potx" or ".potm" => OpenPowerpoint(path),
            _ => new Label { Content = "File not supported." },
        };
    }

    private static Control OpenWord(string path)
    {
        var editor = new SfRichTextBoxAdv
        {
            IsReadOnly = true,
            Background = Brushes.Transparent,
            EnableMiniToolBar = false,
        };

        editor.LoadAsync(path);

        return editor;
    }

    private static Control OpenExcel(string path)
    {
        // Pre-load Syncfusion.Tools.Wpf.dll or it will throw error
        var _ = new ToolsWPFAssembly();

        var sheet = new SfSpreadsheet
        {
            AllowCellContextMenu = false,
            AllowTabItemContextMenu = false,
            Background = Brushes.Transparent,
            DisplayAlerts = false
        };

        sheet.AddGraphicChartCellRenderer(new GraphicChartCellRenderer());

        sheet.Open(path);

        sheet.WorkbookLoaded += (sender, e) =>
        {
            sheet.SuspendFormulaCalculation();

            sheet.Protect(true, true, string.Empty);
            sheet.Workbook.Worksheets.ForEach(s => sheet.ProtectSheet(s, string.Empty));
            sheet.GridCollection.ForEach(kv => kv.Value.ShowHidePopup(false));
        };

        return sheet;
    }

    private static Control OpenPowerpoint(string path)
    {
        var ppt = Presentation.Open(path);
        ppt.ChartToImageConverter = new ChartToImageConverter();

        var settings = new PresentationToPdfConverterSettings
        {
            OptimizeIdenticalImages = true,
            ShowHiddenSlides = true
        };

        var pdf = PresentationToPdfConverter.Convert(ppt, settings);

        var viewer = new PdfViewerControl();

        var tempPdf = new MemoryStream();
        pdf.Save(tempPdf);
        pdf.Close(true);
        pdf.Dispose();
        ppt.Close();
        ppt.Dispose();

        viewer.Dispatcher.BeginInvoke(new Action(() =>
        {
            viewer.LoadPdf(tempPdf);
            tempPdf.Dispose();
        }), DispatcherPriority.Loaded);

        return viewer;
    }
}
