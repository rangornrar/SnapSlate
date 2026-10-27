using System;
using System.Threading.Tasks;
using Microsoft.UI.Windowing;
using Microsoft.Windows.ApplicationModel.Resources;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private readonly ResourceManager _safetyResources = new();

    private string SafetyText(string key)
    {
        var context = _safetyResources.CreateResourceContext();
        context.QualifierValues["Language"] = UseEnglish ? "en-US" : "fr-FR";
        return _safetyResources.MainResourceMap.GetValue("Resources/" + key, context).ValueAsString;
    }

    private void MainWindow_Closing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (_allowClose) return;
        // Closing has no async deferral. Cancel synchronously and prompt on the next UI turn.
        args.Cancel = true;
        if (_isCloseConfirmationOpen || _isSavingProject || _isOpeningProject || _isApplyingProjectState || _isProjectDialogOpen) return;
        _isCloseConfirmationOpen = true;
        if (!DispatcherQueue.TryEnqueue(async () => await ConfirmWindowCloseAsync()))
        {
            _isCloseConfirmationOpen = false;
        }
    }

    private async Task ConfirmWindowCloseAsync()
    {
        try
        {
            // Let an import already in flight finish before inspecting the session.
            // New clipboard imports are paused while this confirmation is pending.
            while (_isProcessingClipboard || _isImportingImage)
            {
                await Task.Delay(25);
            }
            if (await EnsureProjectCanBeReplacedAsync(SafetyText("CloseAction")))
            {
                _allowClose = true;
                DispatcherQueue.TryEnqueue(() => Close());
            }
        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            StatusText.Text = SafetyText("OperationFailed");
        }
        finally
        {
            _isCloseConfirmationOpen = false;
        }
    }
}

