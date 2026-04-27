using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using XamlPath = Microsoft.UI.Xaml.Shapes.Path;
using Windows.Foundation;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Microsoft.Windows.Globalization;
using WinRT.Interop;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private void InitializeShellUi()
    {
        _isInitializingUi = true;
        try
        {
            PopulatePaletteOrder();
            ApplyChromeLabels();
            ConfigureComboBoxes();
            ApplyPreferencesToUi();
            BuildPaletteStrip();
            BuildShadeStrip();
            SetSection(ShellSection.Procedure);
            UpdateExportSummaryList();
        }
        finally
        {
            _isInitializingUi = false;
        }
    }

    private bool UseEnglish =>
        App.Preferences.Language == LanguageChoice.English ||
        (App.Preferences.Language == LanguageChoice.System &&
         CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.Equals("en", StringComparison.OrdinalIgnoreCase));

    private string T(string french, string english) => UseEnglish ? english : french;

    private void PopulatePaletteOrder()
    {
        _paletteOrder.Clear();
        foreach (var key in new[] { "Sunset", "Ember", "Citrus", "Mint", "Lagoon", "Sky", "Ocean", "Rose", "Berry", "Grape" })
        {
            if (_palettes.TryGetValue(key, out var palette))
            {
                _paletteOrder.Add(palette);
            }
        }
    }

    private void ApplyChromeLabels()
    {
        CaptureNavItem.Content = T("Capturer / Aide", "Capture / Help");
        ProcedureNavItem.Content = T("Procédure", "Procedure");
        ExportNavItem.Content = T("Export", "Export");
        SettingsNavItem.Content = T("Réglages", "Settings");

        SetButtonContent(CaptureQuickImportButton, T("Importer", "Import"));
        SetButtonContent(ImportScreenshotButton, T("Importer", "Import"));
        SetButtonContent(ExportCurrentButton, T("Exporter", "Export"));
        SetButtonContent(SaveButton, T("Sauver", "Save"));
        SetButtonContent(CollapseStepsButton, T("Cacher étapes", "Hide steps"));
        SetButtonContent(BrowseExportFolderButton, T("Parcourir", "Browse"));
        SetButtonContent(BrowseDefaultExportFolderButton, T("Parcourir", "Browse"));
        SetButtonContent(UndoButton, T("Annuler", "Undo"));
        SetButtonContent(RedoButton, T("Rétablir", "Redo"));
        SetButtonContent(ZoomOutButton, T("−", "−"));
        SetButtonContent(ZoomResetButton, T("100 %", "100 %"));
        SetButtonContent(ZoomInButton, T("+", "+"));
        ZoomLevelText.Text = "100 %";
        SetButtonContent(ExportPngButton, "PNG");
        SetButtonContent(ExportAllFormatsButton, T("Tout", "All"));
        SetButtonContent(ExportMarkdownButton, "Markdown");
        SetButtonContent(ExportHtmlButton, "HTML");
        SetButtonContent(ExportPdfButton, "PDF");
        SetButtonContent(ExportDocxButton, T("Word", "Word"));
        SetButtonContent(AddStepButton, T("Ajouter", "Add"));
        SetButtonContent(MoveStepUpButton, T("↑", "↑"));
        SetButtonContent(MoveStepDownButton, T("↓", "↓"));

        SetToolButton(CropToolButton, "Crop", T("Crop", "Crop"));
        SetToolButton(TextToolButton, "T", T("Texte", "Text"));
        SetToolButton(StickerToolButton, "1", T("Gommette", "Sticker"));
        SetToolButton(ArrowStraightToolButton, "→", T("Fleche droite", "Straight arrow"));
        SetToolButton(ArrowCurvedToolButton, "↷", T("Fleche courbe", "Curved arrow"));
        SetToolButton(RectangleToolButton, "▭", T("Rectangle", "Rectangle"));
        SetToolButton(EllipseToolButton, "◯", T("Ovale", "Ellipse"));
        SetToolButton(FocusToolButton, "◉", T("Focus", "Focus"));
        SetToolButton(MaskToolButton, "■", T("Masquage", "Mask"));

        SetText(CapturePageTitleText, T("Aide rapide", "Quick help"));
        SetText(CapturePageSummaryText, T("Importez une capture, annotez-la, puis exportez un document propre.", "Import a capture, annotate it, then export a clean document."));
        SetText(CaptureHowTitleText, T("Comment ça marche", "How it works"));
        SetText(CaptureHowBodyText, T("Chaque capture devient une étape de Procédure. Les annotations et l’export restent au même endroit.", "Each capture becomes a Procedure step. Annotations and export stay in one place."));
        SetText(CaptureStepOneTitleText, T("1. Capture", "1. Capture"));
        SetText(CaptureStepOneBodyText, T("Win + Shift + S ou import manuel.", "Win + Shift + S or manual import."));
        SetText(CaptureStepTwoTitleText, T("2. Annoter", "2. Annotate"));
        SetText(CaptureStepTwoBodyText, T("Ajoutez gommettes, flèches, focus et masquage.", "Add stickers, arrows, focus and masking."));
        SetText(CaptureStepThreeTitleText, T("3. Exporter", "3. Export"));
        SetText(CaptureStepThreeBodyText, T("PNG, Markdown, HTML, PDF ou DOCX.", "PNG, Markdown, HTML, PDF or DOCX."));
        SetText(CaptureShortcutsTitleText, T("Raccourcis utiles", "Useful shortcuts"));
        SetText(CaptureShortcutOneText, T("Win + Shift + S : capture Windows", "Win + Shift + S: Windows capture"));
        SetText(CaptureShortcutTwoText, T("Ctrl + O : importer une image", "Ctrl + O: import an image"));
        SetText(CaptureShortcutThreeText, T("Suppr : supprimer l’objet sélectionné", "Delete: delete the selected object"));
        SetText(CaptureStatusTitleText, T("État", "Status"));
        SetText(CaptureStatusBodyText, T("L’import automatique fonctionne tant que SnapSlate est ouvert.", "Automatic import works while SnapSlate is open."));
        SetText(ProcedureTitleLabelText, T("Document", "Document"));
        SetText(ProcedureStateLabelText, T("État", "Status"));
        SetText(CurrentToolLabelText, T("Outil", "Tool"));
        SetText(PaletteSectionTitleText, T("Couleurs", "Colors"));
        SetText(ShadeSectionTitleText, T("Nuances", "Shades"));
        SetText(OpacityLabelText, T("Opacité", "Opacity"));
        SetText(StepsPaneTitleText, T("Étapes", "Steps"));
        CanvasSummaryTitleText.Visibility = Visibility.Collapsed;
        CanvasSummaryText.Visibility = Visibility.Collapsed;
        CanvasFileLabelText.Visibility = Visibility.Collapsed;
        CanvasFileNameText.Visibility = Visibility.Collapsed;
        SetText(CanvasPreviewLabelText, T("Aperçu", "Preview"));
        SetText(StepDetailsTitleText, T("Détails de l’étape", "Step details"));
        SetText(StepTitleLabelText, T("Titre de l’étape", "Step title"));
        SetText(StepNoteLabelText, T("Note de l’étape", "Step note"));
        SetText(TemplateTypeLabelText, T("Type de document", "Document type"));
        SetText(AudienceLabelText, T("Public visé", "Audience"));
        SetText(AuthorLabelText, T("Auteur", "Author"));
        SetText(DocumentVersionLabelText, T("Version", "Version"));
        SetText(AnnotationPanelTitleText, T("Objet sélectionné", "Selected object"));
        SetText(StickerValueLabelText, T("Valeur de la gommette", "Sticker value"));
        SetText(StickerLegendLabelText, T("Légende de la gommette", "Sticker legend"));
        SetText(AnnotationTextLabelText, T("Texte", "Text"));
        SetText(StickerModeLabelText, T("Format de la gommette", "Sticker format"));
        SetText(StrokeThicknessLabelText, T("Épaisseur", "Thickness"));
        SetText(SelectedOpacityLabelText, T("Opacité", "Opacity"));
        SetText(LegendTitleText, T("Légendes des gommettes", "Sticker legends"));
        SetText(ExportPageTitleText, T("Export", "Export"));
        SetText(ExportPageSummaryText, T("PNG pour l’étape courante. Les autres boutons exportent le document au format choisi, ou en lot avec Tout.", "PNG for the current step. The other buttons export the document in the chosen format, or in a batch with All."));
        SetText(ExportTemplateLabelText, T("Type de document", "Document type"));
        SetText(ExportFormatLabelText, T("Format principal", "Main format"));
        SetText(ExportAudienceLabelText, T("Public", "Audience"));
        SetText(ExportAuthorLabelText, T("Auteur", "Author"));
        SetText(ExportVersionLabelText, T("Version", "Version"));
        SetText(ExportFolderLabelText, T("Dossier d’export", "Export folder"));
        SetText(ExportTemplateValueText, T("Type actuel : Manuel utilisateur", "Current type: User manual"));
        SetText(ExportPreviewTitleText, T("Résumé", "Summary"));
        SetText(ExportPreviewSummaryText, T("Tout génère Markdown, HTML, PDF et DOCX dans le même dossier que les images.", "All generates Markdown, HTML, PDF and DOCX in the same folder as the images."));
        SetText(SettingsPageTitleText, T("Réglages", "Settings"));
        SetText(SettingsPageSummaryText, T("Choisissez le thème, la langue et les comportements de capture.", "Choose the theme, language and capture behavior."));
        SetText(AppearanceSettingsTitleText, T("Apparence", "Appearance"));
        SetText(LanguageSettingsTitleText, T("Langue", "Language"));
        SetText(CaptureSettingsTitleText, T("Capture", "Capture"));
        SetText(ExportSettingsTitleText, T("Export par défaut", "Default export"));
        SetText(ThemeSettingLabelText, T("Thème", "Theme"));
        SetText(LanguageSettingLabelText, T("Langue de l’application", "App language"));
        SetText(CaptureSettingLabelText, T("Import automatique", "Automatic import"));
        SetText(DefaultExportFolderLabelText, T("Dossier d’export", "Export folder"));
        SetText(DefaultExportFormatSettingLabelText, T("Format par défaut", "Default format"));
        SetText(SettingsShortcutsTitleText, T("Raccourcis", "Shortcuts"));
        SetText(SettingsShortcutsSummaryText, T("Les raccourcis utiles pour produire une procédure plus vite.", "Useful shortcuts to produce a procedure faster."));
        SetText(ThemeSettingHelpText, T("Le mode Système suit le thème Windows automatiquement.", "System follows Windows automatically."));
        SetText(LanguageSettingHelpText, T("Par défaut, SnapSlate suit la langue de Windows.", "By default, SnapSlate follows the Windows language."));
        SetText(CaptureSettingHelpText, T("Active l’import automatique des captures copiées dans le presse-papiers.", "Enables automatic import of captures copied to the clipboard."));
        SetText(DefaultExportFolderHelpText, T("Choisissez le dossier où SnapSlate créera les exports. Downloads est utilisé par défaut.", "Choose the folder where SnapSlate will create exports. Downloads is used by default."));
        SetText(DefaultExportFormatHelpText, T("Ce format sera présélectionné lors de vos prochains exports.", "This format will be preselected for your next exports."));
        SetText(SettingsRestartHintText, T("Le thème et la langue s’appliquent immédiatement, sans redémarrage.", "Theme and language apply immediately without restarting."));
        AutoImportClipboardCheckBox.Content = T("Importer automatiquement les captures du presse-papiers", "Automatically import clipboard captures");
    }

    private void UpdateExportSummaryList()
    {
        ExportStepsListView.ItemsSource = new[]
        {
            T("PNG de l’étape courante", "Current step PNG"),
            T("Markdown + images", "Markdown + images"),
            T("HTML + images", "HTML + images"),
            T("PDF du document", "Document PDF"),
            T("Word / DOCX", "Word / DOCX")
        };
        SettingsShortcutsListView.ItemsSource = new[]
        {
            T("Win + Shift + S : importer la capture active", "Win + Shift + S: import the current capture"),
            T("Ctrl + O : importer un fichier image", "Ctrl + O: import an image file"),
            T("Suppr : supprimer l’objet sélectionné", "Delete: delete the selected object")
        };
    }

    private string GetTemplateDisplayName(GuideTemplateType type)
    {
        return type switch
        {
            GuideTemplateType.UserManual => T("Manuel utilisateur", "User manual"),
            GuideTemplateType.InternalProcedure => T("Procédure interne", "Internal procedure"),
            GuideTemplateType.SupportPlaybook => T("Support client", "Support playbook"),
            GuideTemplateType.ReleaseNotes => T("Notes de version", "Release notes"),
            _ => T("Manuel utilisateur", "User manual")
        };
    }

    private void UpdateExportTemplateSummary()
    {
        if (ExportTemplateValueText is null)
        {
            return;
        }

        var templateType = _currentDocument?.TemplateType ?? GuideTemplateType.UserManual;
        ExportTemplateValueText.Text = string.Format(
            CultureInfo.CurrentCulture,
            T("Type actuel : {0}", "Current type: {0}"),
            GetTemplateDisplayName(templateType));
    }

    private void ConfigureComboBoxes()
    {
        PopulateComboBox(ThemeComboBox,
            (T("Système", "System"), ThemeChoice.System),
            (T("Clair", "Light"), ThemeChoice.Light),
            (T("Sombre", "Dark"), ThemeChoice.Dark));

        PopulateComboBox(LanguageComboBox,
            (T("Suivre Windows", "Follow Windows"), LanguageChoice.System),
            (T("Français", "French"), LanguageChoice.French),
            (T("English", "English"), LanguageChoice.English));

        PopulateComboBox(DefaultExportFormatComboBox,
            ("PNG", ExportFormatChoice.Png),
            ("Markdown", ExportFormatChoice.Markdown),
            ("HTML", ExportFormatChoice.Html),
            ("PDF", ExportFormatChoice.Pdf),
            (T("Word", "Word"), ExportFormatChoice.Docx));

        PopulateComboBox(ExportFormatComboBox,
            ("PNG", ExportFormatChoice.Png),
            ("Markdown", ExportFormatChoice.Markdown),
            ("HTML", ExportFormatChoice.Html),
            ("PDF", ExportFormatChoice.Pdf),
            (T("Word", "Word"), ExportFormatChoice.Docx));

        PopulateComboBox(TemplateTypeComboBox,
            (T("Manuel utilisateur", "User manual"), GuideTemplateType.UserManual),
            (T("Procédure interne", "Internal procedure"), GuideTemplateType.InternalProcedure),
            (T("Support client", "Support playbook"), GuideTemplateType.SupportPlaybook),
            (T("Notes de version", "Release notes"), GuideTemplateType.ReleaseNotes));

        PopulateComboBox(ExportTemplateComboBox,
            (T("Manuel utilisateur", "User manual"), GuideTemplateType.UserManual),
            (T("Procédure interne", "Internal procedure"), GuideTemplateType.InternalProcedure),
            (T("Support client", "Support playbook"), GuideTemplateType.SupportPlaybook),
            (T("Notes de version", "Release notes"), GuideTemplateType.ReleaseNotes));

        TemplateTypeComboBox.SelectedIndex = 0;
        ExportTemplateComboBox.SelectedIndex = 0;

        PopulateComboBox(StickerModeComboBox,
            (T("Numérique", "Numeric"), 0),
            (T("Alphabétique", "Alphabetic"), 1));
    }

    private void PopulateComboBox<T>(ComboBox comboBox, params (string label, T value)[] items)
    {
        comboBox.Items.Clear();
        foreach (var (label, value) in items)
        {
            comboBox.Items.Add(new ComboBoxItem { Content = label, Tag = value });
        }
    }

    private void ApplyPreferencesToUi()
    {
        RootGrid.RequestedTheme = App.Preferences.GetRequestedTheme();
        ThemeComboBox.SelectedIndex = (int)App.Preferences.Theme;
        LanguageComboBox.SelectedIndex = (int)App.Preferences.Language;
        DefaultExportFormatComboBox.SelectedIndex = (int)App.Preferences.DefaultExportFormat;
        ExportFormatComboBox.SelectedIndex = (int)App.Preferences.DefaultExportFormat;
        AutoImportClipboardCheckBox.IsChecked = App.Preferences.AutoImportClipboard;
        DefaultExportFolderTextBox.Text = App.Preferences.ExportFolder;
        ExportFolderTextBox.Text = App.Preferences.ExportFolder;
        StepsPaneColumn.Width = App.Preferences.CollapseStepsPane ? new GridLength(0) : new GridLength(220);
        InspectorPaneColumn.Width = new GridLength(220);
        CollapseStepsButton.Content = App.Preferences.CollapseStepsPane ? T("Montrer étapes", "Show steps") : T("Cacher étapes", "Hide steps");
    }

    private void BuildPaletteStrip()
    {
        PaletteStrip.Children.Clear();
        _paletteButtons.Clear();

        foreach (var palette in _paletteOrder)
        {
            var button = new Button
            {
                Tag = palette.Key,
                Content = palette.DisplayName,
                MinWidth = 48,
                Padding = new Thickness(5, 2, 5, 2),
                FontSize = 10,
                Background = CreateGradientBrush(palette.Key, 3),
                Foreground = GetBrush("ShellStrongTextBrush"),
                BorderThickness = new Thickness(1)
            };
            button.Click += PaletteButton_Click;
            _paletteButtons[palette.Key] = button;
            ToolTipService.SetToolTip(button, palette.DisplayName);
            AutomationProperties.SetName(button, palette.DisplayName);
            PaletteStrip.Children.Add(button);
        }
    }

    private void BuildShadeStrip()
    {
        ShadeStrip.Children.Clear();

        for (var i = 0; i < 7; i++)
        {
            var button = new Button
            {
                Tag = i,
                Width = 20,
                Height = 20,
                Padding = new Thickness(0),
                Content = (i + 1).ToString(),
                FontSize = 9
            };
            button.Click += ShadeButton_Click;
            var tooltip = string.Format(CultureInfo.CurrentCulture, T("Nuance {0}", "Shade {0}"), i + 1);
            ToolTipService.SetToolTip(button, tooltip);
            AutomationProperties.SetName(button, tooltip);
            ShadeStrip.Children.Add(button);
        }

        RefreshShadeStrip();
    }

    private void RefreshShadeStrip()
    {
        if (_currentDocument is null || !_palettes.TryGetValue(_currentDocument.SelectedPaletteKey, out var palette))
        {
            return;
        }

        foreach (var child in ShadeStrip.Children.OfType<Button>())
        {
            if (child.Tag is not int index)
            {
                continue;
            }

            var shadeIndex = Math.Clamp(index, 0, palette.Shades.Count - 1);
            child.Background = new SolidColorBrush(palette.Shades[shadeIndex]);
            child.BorderBrush = GetBrush(index == _currentDocument.SelectedPaletteShadeIndex ? "ToolSelectedBrush" : "ShellStrokeBrush");
            child.BorderThickness = new Thickness(index == _currentDocument.SelectedPaletteShadeIndex ? 3 : 1);
        }
    }

    private void SetButtonContent(Button button, string content) => button.Content = content;

    private void SetToolButton(Button button, string content, string tooltip)
    {
        button.Content = content;
        ToolTipService.SetToolTip(button, tooltip);
        AutomationProperties.SetName(button, tooltip);
    }

    private void SetText(TextBlock textBlock, string content) => textBlock.Text = content;

    private void SetSection(ShellSection section)
    {
        _currentSection = section;
        CaptureSection.Visibility = section == ShellSection.Capture ? Visibility.Visible : Visibility.Collapsed;
        ProcedureSection.Visibility = section == ShellSection.Procedure ? Visibility.Visible : Visibility.Collapsed;
        ExportSection.Visibility = section == ShellSection.Export ? Visibility.Visible : Visibility.Collapsed;
        SettingsSection.Visibility = section == ShellSection.Settings ? Visibility.Visible : Visibility.Collapsed;
    }

    private void ShellNavigationView_SelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (_isInitializingUi)
        {
            return;
        }

        if (args.SelectedItemContainer is NavigationViewItem item && item.Tag is string tag)
        {
            SetSection(tag switch
            {
                "Capture" => ShellSection.Capture,
                "Export" => ShellSection.Export,
                "Settings" => ShellSection.Settings,
                _ => ShellSection.Procedure
            });
        }
    }

    private void CollapseStepsButton_Click(object sender, RoutedEventArgs e)
    {
        App.Preferences.CollapseStepsPane = !App.Preferences.CollapseStepsPane;
        App.Preferences.Save();
        ApplyPreferencesToUi();
    }

    private void MoveDocumentUpButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: ScreenshotDocument document })
        {
            MoveDocument(document, -1);
        }
    }

    private void MoveDocumentDownButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: ScreenshotDocument document })
        {
            MoveDocument(document, 1);
        }
    }

    private void MoveStepUpButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is not null)
        {
            MoveDocument(_currentDocument, -1);
        }
    }

    private void MoveStepDownButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is not null)
        {
            MoveDocument(_currentDocument, 1);
        }
    }

    private void DocumentTabsListView_DragItemsCompleted(ListViewBase sender, DragItemsCompletedEventArgs args)
    {
        UpdateStatus();
    }

    private void DuplicateDocumentButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: ScreenshotDocument document })
        {
            var duplicate = CloneDocument(document);
            var index = _documents.IndexOf(document);
            _documents.Insert(Math.Max(0, index + 1), duplicate);
            DocumentTabsListView.SelectedItem = duplicate;
        }
    }

    private void UndoButton_Click(object sender, RoutedEventArgs e) => StatusText.Text = T("Annuler n’est pas encore branché.", "Undo is not wired yet.");

    private void RedoButton_Click(object sender, RoutedEventArgs e) => StatusText.Text = T("Rétablir n’est pas encore branché.", "Redo is not wired yet.");

    private void ZoomInButton_Click(object sender, RoutedEventArgs e) => SetZoom(_currentZoom + 0.1);

    private void ZoomOutButton_Click(object sender, RoutedEventArgs e) => SetZoom(_currentZoom - 0.1);

    private void ZoomResetButton_Click(object sender, RoutedEventArgs e) => SetZoom(1.0);

    private void SetZoom(double zoom)
    {
        _currentZoom = Math.Clamp(zoom, 0.5, 2.0);
        SceneScaleTransform.ScaleX = _currentZoom;
        SceneScaleTransform.ScaleY = _currentZoom;
        ZoomLevelText.Text = string.Format(CultureInfo.CurrentCulture, "{0:0} %", _currentZoom * 100);
        UpdateStatus();
    }

    private void DocumentTitleTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        _currentDocument.BaseTitle = textBox.Text.Trim();
        MarkDocumentDirty(_currentDocument);
    }

    private void StepNoteTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        _currentDocument.StepNote = textBox.Text;
        MarkDocumentDirty(_currentDocument);
        UpdateLegendList();
    }

    private void TemplateTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null)
        {
            return;
        }

        if (TemplateTypeComboBox.SelectedItem is ComboBoxItem { Tag: GuideTemplateType type })
        {
            _currentDocument.TemplateType = type;
            MarkDocumentDirty(_currentDocument);
            if (ExportTemplateComboBox.SelectedIndex != TemplateTypeComboBox.SelectedIndex)
            {
                ExportTemplateComboBox.SelectedIndex = TemplateTypeComboBox.SelectedIndex;
            }
            UpdateExportTemplateSummary();
        }
    }

    private void ExportTemplateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null)
        {
            return;
        }

        if (ExportTemplateComboBox.SelectedItem is ComboBoxItem { Tag: GuideTemplateType type })
        {
            _currentDocument.TemplateType = type;
            MarkDocumentDirty(_currentDocument);
            if (TemplateTypeComboBox.SelectedIndex != ExportTemplateComboBox.SelectedIndex)
            {
                TemplateTypeComboBox.SelectedIndex = ExportTemplateComboBox.SelectedIndex;
            }
            UpdateExportTemplateSummary();
        }
    }

    private void AudienceTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        _currentDocument.Audience = textBox.Text;
    }

    private void AuthorTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        _currentDocument.Author = textBox.Text;
    }

    private void DocumentVersionTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        _currentDocument.DocumentVersion = textBox.Text;
    }

    private void ExportAudienceTextBox_TextChanged(object sender, TextChangedEventArgs e) => AudienceTextBox_TextChanged(sender, e);

    private void ExportAuthorTextBox_TextChanged(object sender, TextChangedEventArgs e) => AuthorTextBox_TextChanged(sender, e);

    private void ExportVersionTextBox_TextChanged(object sender, TextChangedEventArgs e) => DocumentVersionTextBox_TextChanged(sender, e);

    private void ExportFolderTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        App.Preferences.ExportFolder = textBox.Text;
        App.Preferences.Save();
        if (!_isApplyingDocumentState)
        {
            DefaultExportFolderTextBox.Text = textBox.Text;
        }
    }

    private void DefaultExportFolderTextBox_TextChanged(object sender, TextChangedEventArgs e) => ExportFolderTextBox_TextChanged(sender, e);

    private async void BrowseExportFolderButton_Click(object sender, RoutedEventArgs e)
    {
        await BrowseForExportFolderAsync();
    }

    private async void BrowseDefaultExportFolderButton_Click(object sender, RoutedEventArgs e)
    {
        await BrowseForExportFolderAsync();
    }

    private async Task BrowseForExportFolderAsync()
    {
        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");
        picker.SuggestedStartLocation = PickerLocationId.Downloads;
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var folder = await picker.PickSingleFolderAsync();
        if (folder is null)
        {
            StatusText.Text = T("Choix de dossier annulé.", "Folder selection cancelled.");
            return;
        }

        App.Preferences.ExportFolder = folder.Path;
        App.Preferences.Save();
        if (DefaultExportFolderTextBox is not null)
        {
            DefaultExportFolderTextBox.Text = folder.Path;
        }

        if (ExportFolderTextBox is not null)
        {
            ExportFolderTextBox.Text = folder.Path;
        }

        StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("Dossier d’export : {0}", "Export folder: {0}"), folder.Path);
    }

    private void DefaultExportFormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || DefaultExportFormatComboBox.SelectedItem is not ComboBoxItem { Tag: ExportFormatChoice format })
        {
            return;
        }

        App.Preferences.DefaultExportFormat = format;
        App.Preferences.Save();
        ExportFormatComboBox.SelectedIndex = DefaultExportFormatComboBox.SelectedIndex;
    }

    private void ExportFormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || ExportFormatComboBox.SelectedItem is not ComboBoxItem { Tag: ExportFormatChoice format })
        {
            return;
        }

        App.Preferences.DefaultExportFormat = format;
        App.Preferences.Save();
        DefaultExportFormatComboBox.SelectedIndex = ExportFormatComboBox.SelectedIndex;
    }

    private void ThemeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || ThemeComboBox.SelectedItem is not ComboBoxItem { Tag: ThemeChoice theme })
        {
            return;
        }

        App.Preferences.Theme = theme;
        App.Preferences.Save();
        RootGrid.RequestedTheme = App.Preferences.GetRequestedTheme();
    }

    private void LanguageComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || LanguageComboBox.SelectedItem is not ComboBoxItem { Tag: LanguageChoice language })
        {
            return;
        }

        App.Preferences.Language = language;
        App.Preferences.Save();
        ApplyChromeLabels();
    }

    private void AutoImportClipboardCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        if (_isInitializingUi)
        {
            return;
        }

        App.Preferences.AutoImportClipboard = AutoImportClipboardCheckBox.IsChecked == true;
        App.Preferences.Save();
    }

    private void OpacitySlider_ValueChanged(object sender, RangeBaseValueChangedEventArgs e)
    {
        if (_currentDocument is null || _isInitializingUi || _isApplyingDocumentState)
        {
            return;
        }

        _currentDocument.DefaultOpacity = OpacitySlider.Value;
    }

    private void SelectedOpacitySlider_ValueChanged(object sender, RangeBaseValueChangedEventArgs e)
    {
        if (_currentDocument is null || _isInitializingUi || _isApplyingDocumentState || _isSyncingAnnotationSelectionUi)
        {
            return;
        }

        if (_selectedAnnotation is not null)
        {
            _selectedAnnotation.Opacity = SelectedOpacitySlider.Value;
            RenderAnnotations(_currentDocument);
            SelectAnnotation(_selectedAnnotation);
            return;
        }

        _currentDocument.DefaultOpacity = SelectedOpacitySlider.Value;
    }

    private async void ShadeButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null || sender is not Button { Tag: int shadeIndex })
        {
            return;
        }

        await ApplyPaletteSelectionAsync(_currentDocument.SelectedPaletteKey, Math.Clamp(shadeIndex, 0, 6));
    }

    private void DuplicateSelectedAnnotationButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null || _selectedAnnotation is null)
        {
            return;
        }

        var duplicate = _selectedAnnotation.Clone();
        duplicate.Id = Guid.NewGuid();
        duplicate.Bounds = new Rect(duplicate.Bounds.X + 12, duplicate.Bounds.Y + 12, duplicate.Bounds.Width, duplicate.Bounds.Height);
        duplicate.StartPoint = new Point(duplicate.StartPoint.X + 12, duplicate.StartPoint.Y + 12);
        duplicate.EndPoint = new Point(duplicate.EndPoint.X + 12, duplicate.EndPoint.Y + 12);
        _currentDocument.Annotations.Add(duplicate);
        _selectedAnnotation = duplicate;
        MarkDocumentDirty(_currentDocument);
        RenderAnnotations(_currentDocument);
        SelectAnnotation(duplicate);
        UpdateAnnotationCount(_currentDocument);
    }

    private async void DeleteSelectedAnnotationButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null || _selectedAnnotation is null)
        {
            return;
        }

        var selected = _selectedAnnotation;
        var shouldRenumber = false;
        if (selected.Kind == AnnotationKind.Sticker)
        {
            var dialog = new ContentDialog
            {
                XamlRoot = RootGrid.XamlRoot,
                Title = T("Supprimer la gommette ?", "Delete the sticker?"),
                Content = T("Voulez-vous renuméroter les gommettes restantes ?", "Do you want to renumber the remaining stickers?"),
                PrimaryButtonText = T("Renuméroter", "Renumber"),
                SecondaryButtonText = T("Supprimer sans renuméroter", "Delete without renumbering"),
                CloseButtonText = T("Annuler", "Cancel"),
                DefaultButton = ContentDialogButton.Primary
            };

            var result = await dialog.ShowAsync();
            if (result == ContentDialogResult.None)
            {
                return;
            }

            shouldRenumber = result == ContentDialogResult.Primary;
        }

        _currentDocument.Annotations.Remove(selected);
        _selectedAnnotation = null;
        if (shouldRenumber)
        {
            RenumberStickers(_currentDocument);
        }

        MarkDocumentDirty(_currentDocument);
        RenderAnnotations(_currentDocument);
        SelectAnnotation(null);
        UpdateAnnotationCount(_currentDocument);
        UpdateLegendList();
    }

    private async void ExportMarkdownButton_Click(object sender, RoutedEventArgs e) => await ExportGuideAsync(ExportFormatChoice.Markdown);

    private async void ExportHtmlButton_Click(object sender, RoutedEventArgs e) => await ExportGuideAsync(ExportFormatChoice.Html);

    private async void ExportPdfButton_Click(object sender, RoutedEventArgs e) => await ExportGuideAsync(ExportFormatChoice.Pdf);

    private async void ExportDocxButton_Click(object sender, RoutedEventArgs e) => await ExportGuideAsync(ExportFormatChoice.Docx);

    private async void ExportAllFormatsButton_Click(object sender, RoutedEventArgs e) =>
        await ExportGuideAsync(ExportFormatChoice.Markdown, ExportFormatChoice.Html, ExportFormatChoice.Pdf, ExportFormatChoice.Docx);

    private void MoveDocument(ScreenshotDocument document, int delta)
    {
        var index = _documents.IndexOf(document);
        if (index < 0)
        {
            return;
        }

        var targetIndex = Math.Clamp(index + delta, 0, _documents.Count - 1);
        if (targetIndex == index)
        {
            return;
        }

        _documents.Move(index, targetIndex);
        DocumentTabsListView.SelectedItem = document;
        UpdateStatus();
    }

    private void SyncDocumentEditors(ScreenshotDocument document)
    {
        DocumentTitleTextBox.Text = document.BaseTitle;
        StepTitleTextBox.Text = document.StepTitle;
        StepNoteTextBox.Text = document.StepNote;
        TemplateTypeComboBox.SelectedIndex = (int)document.TemplateType;
        AudienceTextBox.Text = document.Audience;
        AuthorTextBox.Text = document.Author;
        DocumentVersionTextBox.Text = document.DocumentVersion;
        ExportAudienceTextBox.Text = document.Audience;
        ExportAuthorTextBox.Text = document.Author;
        ExportVersionTextBox.Text = document.DocumentVersion;
        AnnotationTextBox.Text = document.AnnotationText;
        StickerModeComboBox.SelectedIndex = document.StickerModeIndex == 1 ? 1 : 0;
        OpacitySlider.Value = document.DefaultOpacity;
        SelectedOpacitySlider.Value = _selectedAnnotation?.Opacity ?? document.DefaultOpacity;
        ResetStickerNumberOnColorChangeCheckBox.IsChecked = document.ResetStickerNumberOnColorChange;
        ExportTemplateComboBox.SelectedIndex = (int)document.TemplateType;
        DefaultExportFolderTextBox.Text = App.Preferences.ExportFolder;
        ExportFolderTextBox.Text = App.Preferences.ExportFolder;
        DefaultExportFormatComboBox.SelectedIndex = (int)App.Preferences.DefaultExportFormat;
        ExportFormatComboBox.SelectedIndex = (int)App.Preferences.DefaultExportFormat;
        UpdateLegendList();
        RefreshShadeStrip();
        UpdateExportTemplateSummary();
    }

    private void UpdateLegendList()
    {
        if (_currentDocument is null)
        {
            LegendListView.ItemsSource = Array.Empty<string>();
            LegendPanel.Visibility = Visibility.Collapsed;
            return;
        }

        var stickers = _currentDocument.Annotations
            .Where(annotation => annotation.Kind == AnnotationKind.Sticker)
            .Select(annotation =>
            {
                var legend = string.IsNullOrWhiteSpace(annotation.LegendText)
                    ? T("À compléter", "To fill in")
                    : annotation.LegendText;
                return $"{annotation.Text} : {legend}";
            })
            .ToArray();

        LegendListView.ItemsSource = stickers;
        LegendPanel.Visibility = stickers.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void SelectAnnotation(AnnotationModel? annotation)
    {
        _selectedAnnotation = annotation;
        if (_currentDocument is null)
        {
            return;
        }

        _isSyncingAnnotationSelectionUi = true;
        try
        {
            if (annotation is null)
            {
                StepDetailsPanel.Visibility = Visibility.Visible;
                AnnotationContextPanel.Visibility = Visibility.Collapsed;
                SelectionOutline.Visibility = Visibility.Collapsed;
                SelectedAnnotationSummaryText.Text = T("Aucun objet sélectionné.", "No object selected.");
                SelectedOpacitySlider.Value = _currentDocument.DefaultOpacity;
                return;
            }

            StepDetailsPanel.Visibility = Visibility.Collapsed;
            AnnotationContextPanel.Visibility = Visibility.Visible;
            AnnotationPanelTitleText.Text = annotation.Kind switch
            {
                AnnotationKind.Sticker => T("Gommette", "Sticker"),
                AnnotationKind.Text => T("Texte", "Text"),
                AnnotationKind.Rectangle => T("Rectangle", "Rectangle"),
                AnnotationKind.Ellipse => T("Ovale", "Ellipse"),
                AnnotationKind.ArrowStraight => T("Flèche droite", "Straight arrow"),
                AnnotationKind.ArrowCurved => T("Flèche courbe", "Curved arrow"),
                AnnotationKind.Focus => T("Focus", "Focus"),
                AnnotationKind.Mask => T("Masquage", "Mask"),
                _ => T("Objet", "Object")
            };

            var isSticker = annotation.Kind == AnnotationKind.Sticker;
            var isText = annotation.Kind == AnnotationKind.Text;
            var isStrokeBased = annotation.Kind is AnnotationKind.Rectangle or AnnotationKind.Ellipse or AnnotationKind.ArrowStraight or AnnotationKind.ArrowCurved or AnnotationKind.Focus or AnnotationKind.Mask;

            StickerValueLabelText.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            StickerValueText.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            StickerLegendLabelText.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            StickerLegendTextBox.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            AnnotationTextLabelText.Visibility = isText ? Visibility.Visible : Visibility.Collapsed;
            AnnotationTextBox.Visibility = isText ? Visibility.Visible : Visibility.Collapsed;
            StickerModeLabelText.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            StickerModeComboBox.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            ResetStickerNumberOnColorChangeCheckBox.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            StrokeThicknessSlider.Visibility = isStrokeBased ? Visibility.Visible : Visibility.Collapsed;
            StrokeSummaryText.Visibility = isStrokeBased ? Visibility.Visible : Visibility.Collapsed;
            CropStateText.Visibility = (_currentDocument.PendingCropRect is not null || _currentDocument.AppliedCropRect is not null) ? Visibility.Visible : Visibility.Collapsed;
            ApplyCropButton.Visibility = _currentDocument.PendingCropRect is not null ? Visibility.Visible : Visibility.Collapsed;
            ResetCropButton.Visibility = (_currentDocument.PendingCropRect is not null || _currentDocument.AppliedCropRect is not null) ? Visibility.Visible : Visibility.Collapsed;

            StickerValueLabelText.Text = T("Valeur de la gommette", "Sticker value");
            StickerValueText.Text = annotation.Text;
            StickerLegendLabelText.Text = T("Légende de la gommette", "Sticker legend");
            StickerLegendTextBox.Text = annotation.LegendText;
            StickerLegendTextBox.PlaceholderText = T("Décrivez ce que pointe cette gommette", "Describe what this sticker points to");
            AnnotationTextLabelText.Text = T("Texte", "Text");
            AnnotationTextBox.Text = annotation.Text;
            AnnotationTextBox.PlaceholderText = T("Saisissez le texte", "Enter text");
            SelectedAnnotationSummaryText.Text = annotation.Kind switch
            {
                AnnotationKind.Sticker or AnnotationKind.Text => $"{annotation.Text} · {annotation.Opacity:P0}",
                AnnotationKind.Rectangle or AnnotationKind.Ellipse => $"{annotation.Bounds.Width:0} x {annotation.Bounds.Height:0} · {annotation.Opacity:P0}",
                _ => annotation.Opacity.ToString("P0", CultureInfo.CurrentCulture)
            };

            SelectedOpacitySlider.Value = annotation.Opacity;
            StrokeThicknessSlider.Value = annotation.StrokeThickness;
            AnnotationTextBox.Text = isText ? annotation.Text : string.Empty;
            UpdateSelectionOutline(annotation);
        }
        finally
        {
            _isSyncingAnnotationSelectionUi = false;
        }
    }

    private void UpdateSelectionOutline(AnnotationModel annotation)
    {
        var bounds = GetAnnotationVisualBounds(annotation, _draggingAnnotationElement);

        if (bounds.Width <= 0 || bounds.Height <= 0)
        {
            SelectionOutline.Visibility = Visibility.Collapsed;
            return;
        }

        SelectionOutline.Visibility = Visibility.Visible;
        SelectionOutline.Width = bounds.Width + 12;
        SelectionOutline.Height = bounds.Height + 12;
        Canvas.SetLeft(SelectionOutline, Math.Max(0, bounds.X - 6));
        Canvas.SetTop(SelectionOutline, Math.Max(0, bounds.Y - 6));
    }

    private void AnnotationElement_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (sender is FrameworkElement { Tag: AnnotationModel annotation } element)
        {
            if (e.GetCurrentPoint(element).Properties.IsRightButtonPressed)
            {
                ClearSelection();
                e.Handled = true;
                return;
            }

            SelectAnnotation(annotation);
            BeginAnnotationDrag(element, annotation, e);
            e.Handled = true;
        }
    }

    private void AnnotationElement_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDraggingAnnotation || sender is not FrameworkElement element || !ReferenceEquals(element, _draggingAnnotationElement))
        {
            return;
        }

        UpdateDraggedAnnotation(e.GetCurrentPoint(InteractionSurface).Position);
        e.Handled = true;
    }

    private void AnnotationElement_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDraggingAnnotation || sender is not FrameworkElement element || !ReferenceEquals(element, _draggingAnnotationElement))
        {
            return;
        }

        EndAnnotationDrag(e);
        e.Handled = true;
    }

    private void AnnotationElement_PointerCaptureLost(object sender, PointerRoutedEventArgs e)
    {
        if (_isDraggingAnnotation && sender is FrameworkElement element && ReferenceEquals(element, _draggingAnnotationElement))
        {
            EndAnnotationDrag(null);
        }
    }

    private void BeginAnnotationDrag(FrameworkElement element, AnnotationModel annotation, PointerRoutedEventArgs e)
    {
        if (_currentDocument is null)
        {
            return;
        }

        _draggingAnnotation = annotation;
        _draggingAnnotationElement = element;
        _isDraggingAnnotation = true;
        _annotationDragStart = e.GetCurrentPoint(InteractionSurface).Position;
        _annotationDragStartBounds = GetAnnotationVisualBounds(annotation, element);
        _annotationDragStartPoint = annotation.StartPoint;
        _annotationDragStartEndPoint = annotation.EndPoint;
        element.CapturePointer(e.Pointer);
    }

    private void UpdateDraggedAnnotation(Point currentPoint)
    {
        if (_draggingAnnotation is null || _draggingAnnotationElement is null)
        {
            return;
        }

        var delta = new Point(currentPoint.X - _annotationDragStart.X, currentPoint.Y - _annotationDragStart.Y);
        switch (_draggingAnnotation.Kind)
        {
            case AnnotationKind.ArrowStraight:
            case AnnotationKind.ArrowCurved:
                _draggingAnnotation.StartPoint = new Point(_annotationDragStartPoint.X + delta.X, _annotationDragStartPoint.Y + delta.Y);
                _draggingAnnotation.EndPoint = new Point(_annotationDragStartEndPoint.X + delta.X, _annotationDragStartEndPoint.Y + delta.Y);
                if (_draggingAnnotationElement is XamlPath path)
                {
                    path.Data = _draggingAnnotation.Kind == AnnotationKind.ArrowStraight
                        ? BuildStraightArrowGeometry(_draggingAnnotation.StartPoint, _draggingAnnotation.EndPoint, _draggingAnnotation.StrokeThickness)
                        : BuildCurvedArrowGeometry(_draggingAnnotation.StartPoint, _draggingAnnotation.EndPoint, _draggingAnnotation.StrokeThickness);
                }
                break;
            default:
                _draggingAnnotation.Bounds = new Rect(
                    _annotationDragStartBounds.X + delta.X,
                    _annotationDragStartBounds.Y + delta.Y,
                    Math.Max(1, _annotationDragStartBounds.Width),
                    Math.Max(1, _annotationDragStartBounds.Height));
                SetCanvasPosition(_draggingAnnotationElement, _draggingAnnotation.Bounds);
                break;
        }

        UpdateSelectionOutline(_draggingAnnotation);
    }

    private void EndAnnotationDrag(PointerRoutedEventArgs? args)
    {
        if (!_isDraggingAnnotation)
        {
            return;
        }

        if (_draggingAnnotationElement is not null && args is not null)
        {
            _draggingAnnotationElement.ReleasePointerCapture(args.Pointer);
        }

        _isDraggingAnnotation = false;
        _draggingAnnotation = null;
        _draggingAnnotationElement = null;

        if (_currentDocument is null)
        {
            return;
        }

        var document = _currentDocument;
        var selectedAnnotation = _selectedAnnotation;
        MarkDocumentDirty(_currentDocument);
        UpdateAnnotationCount(_currentDocument);
        UpdateStatus();

        _ = DispatcherQueue.TryEnqueue(() =>
        {
            if (_currentDocument is null || !ReferenceEquals(_currentDocument, document))
            {
                return;
            }

            RenderAnnotations(document);
            if (selectedAnnotation is not null && ReferenceEquals(_selectedAnnotation, selectedAnnotation))
            {
                SelectAnnotation(selectedAnnotation);
            }
        });
    }

    private Rect GetAnnotationVisualBounds(AnnotationModel annotation, FrameworkElement? element = null)
    {
        return annotation.Kind switch
        {
            AnnotationKind.ArrowStraight or AnnotationKind.ArrowCurved => NormalizeRect(annotation.StartPoint, annotation.EndPoint),
            AnnotationKind.Text => new Rect(
                annotation.Bounds.X,
                annotation.Bounds.Y,
                Math.Max(220, Math.Max(annotation.Bounds.Width, element?.ActualWidth ?? annotation.Bounds.Width)),
                Math.Max(60, Math.Max(annotation.Bounds.Height, element?.ActualHeight ?? annotation.Bounds.Height))),
            AnnotationKind.Sticker => new Rect(
                annotation.Bounds.X,
                annotation.Bounds.Y,
                Math.Max(StickerSize, annotation.Bounds.Width),
                Math.Max(StickerSize, annotation.Bounds.Height)),
            _ => annotation.Bounds
        };
    }

    private void RenumberStickers(ScreenshotDocument document)
    {
        var index = 1;
        foreach (var sticker in document.Annotations.Where(annotation => annotation.Kind == AnnotationKind.Sticker))
        {
            sticker.StickerIndex = index;
            sticker.Text = GetStickerModeIndex(document) switch
            {
                1 => ToAlphabetic(index),
                _ => index.ToString()
            };
            index++;
        }

        document.NextStickerIndex = index;
    }

    private ScreenshotDocument CloneDocument(ScreenshotDocument source)
    {
        var duplicate = new ScreenshotDocument
        {
            BaseTitle = $"{source.BaseTitle} - copie",
            StepNote = source.StepNote,
            Origin = source.Origin,
            OriginLabel = source.OriginLabel,
            SourceLabel = source.SourceLabel,
            FileNameLabel = source.FileNameLabel,
            SourcePixelWidth = source.SourcePixelWidth,
            SourcePixelHeight = source.SourcePixelHeight,
            ImageBytes = source.ImageBytes?.ToArray(),
            ThumbnailSource = source.ThumbnailSource,
            SelectedPaletteKey = source.SelectedPaletteKey,
            SelectedPaletteShadeIndex = source.SelectedPaletteShadeIndex,
            PaletteDisplayName = source.PaletteDisplayName,
            AnnotationText = source.AnnotationText,
            TemplateType = source.TemplateType,
            Audience = source.Audience,
            Author = source.Author,
            DocumentVersion = source.DocumentVersion,
            StickerModeIndex = source.StickerModeIndex,
            StrokeThickness = source.StrokeThickness,
            DefaultOpacity = source.DefaultOpacity,
            NextStickerIndex = source.NextStickerIndex,
            ResetStickerNumberOnColorChange = source.ResetStickerNumberOnColorChange,
            PendingCropRect = source.PendingCropRect,
            AppliedCropRect = source.AppliedCropRect,
            IsDirty = true
        };

        foreach (var annotation in source.Annotations)
        {
            duplicate.Annotations.Add(annotation.Clone());
        }

        return duplicate;
    }

    private async Task ExportGuideAsync(params ExportFormatChoice[] formats)
    {
        if (_documents.Count == 0)
        {
            StatusText.Text = T("Aucune étape à exporter.", "No steps to export.");
            return;
        }

        var requestedFormats = formats is { Length: > 0 }
            ? formats.Distinct().ToArray()
            : new[] { ExportFormatChoice.Markdown, ExportFormatChoice.Html, ExportFormatChoice.Pdf, ExportFormatChoice.Docx };

        var rootFolder = App.Preferences.ExportFolder;
        if (string.IsNullOrWhiteSpace(rootFolder))
        {
            rootFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SnapSlate-Exports");
        }

        Directory.CreateDirectory(rootFolder);
        var packageFolder = Path.Combine(rootFolder, $"SnapSlate-{DateTime.Now:yyyyMMdd-HHmmss}");
        var imagesFolder = Path.Combine(packageFolder, "images");
        Directory.CreateDirectory(imagesFolder);

        var exportSteps = new List<GuideExportStep>(_documents.Count);
        var generatedFormats = new List<string>();
        if (requestedFormats.Contains(ExportFormatChoice.Png))
        {
            generatedFormats.Add("PNG");
        }

        for (var i = 0; i < _documents.Count; i++)
        {
            var document = _documents[i];
            var imagePath = Path.Combine(imagesFolder, $"step-{i + 1:00}.png");
            await Task.Run(() => GuideImageRenderer.SaveDocumentImage(document, _palettes, imagePath));

            exportSteps.Add(new GuideExportStep(
                i + 1,
                document.BaseTitle,
                document.StepNote,
                document.OriginLabel,
                document.SourceLabel,
                imagePath,
                document.Annotations
                    .Where(annotation => annotation.Kind == AnnotationKind.Sticker)
                    .Select(annotation => new GuideExportLegendItem(
                        annotation.Text,
                        string.IsNullOrWhiteSpace(annotation.LegendText) ? T("À compléter", "To fill in") : annotation.LegendText))
                    .ToArray()));
        }

        var sourceDocument = _currentDocument ?? _documents[0];
        var manifest = new GuideExportManifest(
            sourceDocument.BaseTitle,
            GetTemplateLabel(sourceDocument.TemplateType),
            sourceDocument.Audience,
            sourceDocument.Author,
            sourceDocument.DocumentVersion,
            DateTimeOffset.Now,
            packageFolder,
            exportSteps);

        if (requestedFormats.Contains(ExportFormatChoice.Markdown))
        {
            await File.WriteAllTextAsync(Path.Combine(packageFolder, "guide.md"), BuildMarkdown(manifest), Encoding.UTF8);
            generatedFormats.Add("Markdown");
        }

        if (requestedFormats.Contains(ExportFormatChoice.Html))
        {
            await File.WriteAllTextAsync(Path.Combine(packageFolder, "guide.html"), BuildHtml(manifest), Encoding.UTF8);
            generatedFormats.Add("HTML");
        }

        if (requestedFormats.Contains(ExportFormatChoice.Pdf))
        {
            await GuideExporter.ExportPdfAsync(manifest, Path.Combine(packageFolder, "guide.pdf"));
            generatedFormats.Add("PDF");
        }

        if (requestedFormats.Contains(ExportFormatChoice.Docx))
        {
            await GuideExporter.ExportDocxAsync(manifest, Path.Combine(packageFolder, "guide.docx"));
            generatedFormats.Add("DOCX");
        }

        var statusMessage = string.Format(
            CultureInfo.CurrentCulture,
            T("Export terminé : {0} dans {1}", "Export complete: {0} in {1}"),
            string.Join(", ", generatedFormats.Count > 0 ? generatedFormats : new[] { "PNG" }),
            packageFolder);

        if (!string.IsNullOrWhiteSpace(statusMessage))
        {
            StatusText.Text = statusMessage;
        }
    }

    private string GetTemplateLabel(GuideTemplateType templateType)
    {
        return templateType switch
        {
            GuideTemplateType.InternalProcedure => T("Procédure interne", "Internal procedure"),
            GuideTemplateType.SupportPlaybook => T("Support client", "Support playbook"),
            GuideTemplateType.ReleaseNotes => T("Notes de version", "Release notes"),
            _ => T("Manuel utilisateur", "User manual")
        };
    }

    private static string EscapeHtml(string value)
    {
        return System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
    }

    private string BuildMarkdown(GuideExportManifest manifest)
    {
        var markdown = new StringBuilder();
        markdown.AppendLine($"# {manifest.Title}");
        markdown.AppendLine();
        markdown.AppendLine($"- {T("Type de document", "Document type")} : {manifest.TemplateName}");
        markdown.AppendLine($"- {T("Public", "Audience")} : {manifest.Audience}");
        markdown.AppendLine($"- {T("Auteur", "Author")} : {manifest.Author}");
        markdown.AppendLine($"- {T("Version", "Version")} : {manifest.Version}");
        markdown.AppendLine();

        foreach (var step in manifest.Steps)
        {
            markdown.AppendLine($"## {step.Index:00}. {step.Title}");
            markdown.AppendLine();

            if (!string.IsNullOrWhiteSpace(step.Note))
            {
                markdown.AppendLine(step.Note);
                markdown.AppendLine();
            }

            markdown.AppendLine($"![](images/{Path.GetFileName(step.ImagePath)})");
            markdown.AppendLine();

            if (step.LegendItems.Count > 0)
            {
                markdown.AppendLine(T("Légende", "Legend"));
                foreach (var legend in step.LegendItems)
                {
                    markdown.AppendLine($"- {legend.StickerLabel} : {legend.Description}");
                }

                markdown.AppendLine();
            }

            markdown.AppendLine($"_{step.SourceLabel}_");
            markdown.AppendLine();
        }

        return markdown.ToString();
    }

    private string BuildHtml(GuideExportManifest manifest)
    {
        var html = new StringBuilder();
        html.AppendLine("<!doctype html>");
        html.AppendLine($"<html lang=\"{(UseEnglish ? "en" : "fr")}\">");
        html.AppendLine("<head>");
        html.AppendLine("<meta charset=\"utf-8\">");
        html.AppendLine($"<title>{EscapeHtml(manifest.Title)}</title>");
        html.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
        html.AppendLine("<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:1100px;margin:0 auto;padding:32px;color:#1f2937;line-height:1.5}h1,h2{color:#1e3a8a}img{max-width:100%;height:auto;border:1px solid #d1d5db;border-radius:14px}.meta{color:#4b5563}.step{margin:28px 0;padding-bottom:20px;border-bottom:1px solid #e5e7eb}.legend{margin-top:12px;padding:12px 16px;background:#f8fafc;border-radius:12px}.legend ul{margin:8px 0 0 20px}</style>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");
        html.AppendLine($"<h1>{EscapeHtml(manifest.Title)}</h1>");
        html.AppendLine($"<p class=\"meta\">{EscapeHtml(manifest.TemplateName)} · {EscapeHtml(manifest.Audience)} · {EscapeHtml(manifest.Version)}</p>");
        html.AppendLine($"<p class=\"meta\">{EscapeHtml(T("Auteur", "Author"))} : {EscapeHtml(manifest.Author)} · {EscapeHtml(manifest.GeneratedAt.ToString("g", CultureInfo.CurrentCulture))}</p>");

        foreach (var step in manifest.Steps)
        {
            html.AppendLine("<section class=\"step\">");
            html.AppendLine($"<h2>{EscapeHtml($"{step.Index:00}. {step.Title}")}</h2>");
            if (!string.IsNullOrWhiteSpace(step.Note))
            {
                html.AppendLine($"<p>{EscapeHtml(step.Note)}</p>");
            }

            html.AppendLine($"<img src=\"images/{EscapeHtml(Path.GetFileName(step.ImagePath))}\" alt=\"{EscapeHtml(step.Title)}\">");

            if (step.LegendItems.Count > 0)
            {
                html.AppendLine("<div class=\"legend\">");
                html.AppendLine($"<strong>{EscapeHtml(T("Légende", "Legend"))}</strong>");
                html.AppendLine("<ul>");
                foreach (var legend in step.LegendItems)
                {
                    html.AppendLine($"<li><strong>{EscapeHtml(legend.StickerLabel)}</strong> : {EscapeHtml(legend.Description)}</li>");
                }
                html.AppendLine("</ul>");
                html.AppendLine("</div>");
            }

            html.AppendLine($"<p class=\"meta\">{EscapeHtml(step.SourceLabel)}</p>");
            html.AppendLine("</section>");
        }

        html.AppendLine("</body></html>");
        return html.ToString();
    }

}
