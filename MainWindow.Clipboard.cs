using System;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Input;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Streams;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private const string ClipboardImageSourceFormat = "SnapSlate.AnnotatedImage.Source";
    private bool _isCopyingImage;

    private async Task CopyCurrentImageAsync()
    {
        if (_showWelcome || _currentDocument is null || _isCopyingImage || _editorClosed) return;
        _isCopyingImage = true;
        CopyImageButton.IsEnabled = CopyImageMenuItem.IsEnabled = false;
        try
        {
            CommitPendingTextEdits();
            StopInlineTextEdit();
            var snapshot = _currentDocument.CloneForHistory();
            StatusText.Text = SafetyText("ClipboardCopying");
            var bytes = await Task.Run(() => GuideImageRenderer.RenderPng(snapshot, _palettes));
            if (_editorClosed) return;

            using var stream = new InMemoryRandomAccessStream();
            using (var writer = new DataWriter(stream))
            {
                writer.WriteBytes(bytes);
                await writer.StoreAsync();
                writer.DetachStream();
            }
            stream.Seek(0);
            var content = new DataPackage { RequestedOperation = DataPackageOperation.Copy };
            content.SetBitmap(RandomAccessStreamReference.CreateFromStream(stream));
            content.SetData(ClipboardImageSourceFormat, "SnapSlate");
            for (var attempt = 0; attempt < 3; attempt++)
            {
                if (_editorClosed) return;
                if (Clipboard.SetContentWithOptions(content, null))
                {
                    // Persist the rendered image before releasing its stream or closing the app.
                    Clipboard.Flush();
                    StatusText.Text = SafetyText("ClipboardCopySucceeded");
                    return;
                }
                if (attempt < 2) await Task.Delay(100);
            }
            StatusText.Text = SafetyText("ClipboardCopyFailed");
        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            if (!_editorClosed) StatusText.Text = SafetyText("ClipboardCopyFailed");
        }
        finally
        {
            _isCopyingImage = false;
            if (!_editorClosed) CopyImageButton.IsEnabled = CopyImageMenuItem.IsEnabled = true;
        }
    }

    private async void CopyImageButton_Click(object sender, RoutedEventArgs e) => await CopyCurrentImageAsync();

    private async void CopyImageKeyboardAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        args.Handled = true;
        await CopyCurrentImageAsync();
    }
}
