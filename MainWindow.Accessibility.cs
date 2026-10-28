using System;
using System.Linq;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private bool _isSyncingObjectList;
    private string _objectListLanguage = string.Empty;
    private SnapSlateProjectState? _annotationMoveSnapshot;

    private void RefreshAnnotationObjects()
    {
        _isSyncingObjectList = true;
        try
        {
            var annotations = _currentDocument?.Annotations.ToArray() ?? [];
            // Keep the realized containers and keyboard focus when only geometry changes.
            if (_objectListLanguage != LocalizedText.CurrentLanguage
                || AnnotationObjectsList.ItemsSource is not AnnotationModel[] current || !current.SequenceEqual(annotations))
                AnnotationObjectsList.ItemsSource = annotations;
            _objectListLanguage = LocalizedText.CurrentLanguage;
            AnnotationObjectsList.SelectedItem = _selectedAnnotation;
            AnnotationObjectsExpander.Header = SafetyText("AnnotationObjects");
            AutomationProperties.SetName(AnnotationObjectsList, SafetyText("AnnotationObjects"));
            ObjectMoveHelpText.Text = SafetyText("ObjectMoveHelp");
            foreach (var (button, key) in new[] { (MoveObjectLeftButton, "MoveLeft"), (MoveObjectRightButton, "MoveRight"), (MoveObjectUpButton, "MoveUp"), (MoveObjectDownButton, "MoveDown") })
                AutomationProperties.SetName(button, SafetyText(key));
            foreach (var document in _documents) document.NotifyLanguageChanged();
        }
        finally { _isSyncingObjectList = false; }
    }

    private void AnnotationObjectsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_isSyncingObjectList && AnnotationObjectsList.SelectedItem is AnnotationModel annotation
            && !ReferenceEquals(annotation, _selectedAnnotation) && _currentDocument?.Annotations.Contains(annotation) == true)
            SelectAnnotation(annotation);
    }

    private void MoveAnnotationButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null || _selectedAnnotation is null || sender is not Button { Tag: string direction }) return;
        PushUndoState(CaptureProjectState());
        var (x, y) = direction switch { "Left" => (-1, 0), "Right" => (1, 0), "Up" => (0, -1), _ => (0, 1) };
        AnnotationMovement.Translate(_selectedAnnotation, x, y);
        MarkDocumentDirty(_currentDocument);
        RenderAnnotations(_currentDocument);
        UpdateSelectionOutline(_selectedAnnotation);
    }

    private void ApplyAccessibleFieldLabels()
    {
        foreach (var (field, label) in new (DependencyObject, UIElement)[]
        {
            (StepTitleTextBox, StepTitleLabelText), (StepNoteTextBox, StepNoteLabelText),
            (AudienceTextBox, AudienceLabelText), (AuthorTextBox, AuthorLabelText),
            (DocumentVersionTextBox, DocumentVersionLabelText), (AnnotationTextBox, AnnotationTextLabelText),
            (StickerLegendTextBox, StickerLegendLabelText), (TemplateTypeComboBox, TemplateTypeLabelText),
            (ExportTemplateComboBox, ExportTemplateLabelText), (ExportFormatComboBox, ExportFormatLabelText),
            (ExportAudienceTextBox, ExportAudienceLabelText), (ExportAuthorTextBox, ExportAuthorLabelText),
            (ExportVersionTextBox, ExportVersionLabelText), (ExportFolderTextBox, ExportFolderLabelText),
            (ThemeComboBox, ThemeSettingLabelText), (LanguageComboBox, LanguageSettingLabelText),
            (NotionParentPageIdTextBox, NotionParentPageLabelText), (NotionTokenPasswordBox, NotionTokenLabelText),
            (ConfluenceBaseUrlTextBox, ConfluenceBaseUrlLabelText), (ConfluenceSpaceKeyTextBox, ConfluenceSpaceKeyLabelText),
            (ConfluenceParentPageIdTextBox, ConfluenceParentPageLabelText), (ConfluenceEmailTextBox, ConfluenceEmailLabelText),
            (ConfluenceTokenPasswordBox, ConfluenceTokenLabelText), (SharePointSyncFolderTextBox, SharePointFolderLabelText)
        }) AutomationProperties.SetLabeledBy(field, label);
    }
}
