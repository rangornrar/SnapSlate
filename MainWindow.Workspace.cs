using System;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private bool _showStepsPane;
    private bool _inspectorOpenedManually;
    private bool _autoFitPreview = true;
    private bool _workspaceFitQueued;

    private void ApplyWorkspaceChrome()
    {
        void Icon(Button button, string glyph, string label)
        {
            button.Content = new FontIcon { Glyph = glyph, FontSize = 16 };
            button.MinWidth = button.MinHeight = 32;
            button.Padding = new Thickness(8, 4, 8, 4);
            SetButtonTooltip(button, label);
        }

        Icon(SaveProjectButton, "\uE74E", SafetyText("UiSaveTheProject") + " (Ctrl+S)");
        Icon(UndoButton, "\uE7A7", SafetyText("UiUndoTheLastAction") + " (Ctrl+Z)");
        Icon(RedoButton, "\uE7A6", SafetyText("UiRedoTheLastAction") + " (Ctrl+Y)");
        Icon(DuplicateSelectedButton, "\uE8C8", SafetyText("UiDuplicateSelectedObject"));
        Icon(DeleteSelectedButton, "\uE74D", SafetyText("UiDeleteTheSelectedObject"));
        Icon(SettingsMenuButton, "\uE713", SafetyText("UiOpenSettings"));
        ImportScreenshotButton.Content = SafetyText("WorkspaceImport");
        ExportCurrentButton.Content = SafetyText("WorkspaceExportPng");
        ExportCurrentButton.Background = GetBrush("PrimaryActionBrush");
        ExportCurrentButton.Foreground = GetBrush("PrimaryActionTextBrush");
        ExportCurrentButton.Padding = new Thickness(12, 6, 12, 6);
        ToolbarStepsButton.Content = SafetyText("UiSteps");
        InspectorPaneToggleButton.Content = SafetyText("WorkspaceProperties");
        SetButtonTooltip(InspectorPaneToggleButton, SafetyText("UiShowOrHideTheInspector"));
        AutomationProperties.SetName(DocumentTitleTextBox, SafetyText("WorkspaceGuideTitle"));
        AutomationProperties.SetName(ToolbarDocumentComboBox, SafetyText("WorkspaceCapture"));
        AutomationProperties.SetName(ToolbarPaletteComboBox, SafetyText("UiColor"));
        AutomationProperties.SetName(ToolbarOpacitySlider, SafetyText("UiOpacity"));
        UpdateWorkspaceZoomLabel();
        UpdateWorkspaceContext();
    }

    private void UpdateWorkspaceZoomLabel()
    {
        ToolbarZoomResetButton.Content = ZoomLevelText.Text;
        SetButtonTooltip(ToolbarZoomResetButton, ZoomLevelText.Text + " — " + SafetyText("UiResetTo100"));
    }

    private void UpdateWorkspaceContext()
    {
        if (_isDraggingAnnotation) return;
        var crop = _currentTool == EditorTool.Crop || _currentDocument?.PendingCropRect is not null;
        var mask = _selectedAnnotation?.Kind == AnnotationKind.Mask || (_selectedAnnotation is null && _currentTool == EditorTool.Mask);
        WorkspaceContextBar.Visibility = _selectedAnnotation is not null || _currentTool != EditorTool.Select || crop
            ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceBrushControls.Visibility = crop || mask ? Visibility.Collapsed : Visibility.Visible;
        WorkspaceCropControls.Visibility = crop ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceObjectActions.Visibility = _selectedAnnotation is not null && !crop ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceStrokeControls.Visibility = _selectedAnnotation?.Kind is AnnotationKind.Rectangle or AnnotationKind.Ellipse or AnnotationKind.ArrowStraight or AnnotationKind.ArrowCurved or AnnotationKind.Focus
            || (_selectedAnnotation is null && _currentTool is EditorTool.Rectangle or EditorTool.Ellipse or EditorTool.ArrowStraight or EditorTool.ArrowCurved or EditorTool.Focus) ? Visibility.Visible : Visibility.Collapsed;
        StrokeThicknessSlider.Visibility = WorkspaceStrokeControls.Visibility;
        ToolbarCurrentToolText.Text = mask ? SafetyText("WorkspaceOpaqueMask")
            : _selectedAnnotation is not null ? SafetyText("Annotation" + _selectedAnnotation.Kind) : SafetyText(_toolLabels[_currentTool]);
        StepDetailsTitleText.Text = _selectedAnnotation is null ? SafetyText("WorkspaceProperties") : ToolbarCurrentToolText.Text;
    }

    private void QueueWorkspaceFit()
    {
        if (!_autoFitPreview || _workspaceFitQueued) return;
        _workspaceFitQueued = true;
        if (!DispatcherQueue.TryEnqueue(() =>
        {
            try
            {
                if (!_editorClosed && !_isDraggingAnnotation && _autoFitPreview && _currentDocument is not null) FitPreviewToViewport();
            }
            finally { _workspaceFitQueued = false; }
        })) _workspaceFitQueued = false;
    }

    private void SetManualZoom(double zoom)
    {
        _autoFitPreview = false;
        SetZoom(zoom);
    }
}
