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
using Microsoft.UI;
using XamlEllipse = Microsoft.UI.Xaml.Shapes.Ellipse;
using XamlLine = Microsoft.UI.Xaml.Shapes.Line;
using XamlPolygon = Microsoft.UI.Xaml.Shapes.Polygon;
using XamlRectangle = Microsoft.UI.Xaml.Shapes.Rectangle;
using XamlPath = Microsoft.UI.Xaml.Shapes.Path;
using Windows.Foundation;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Microsoft.Windows.Globalization;
using WinRT.Interop;
using Windows.UI;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private static Brush? _iconBrushOverride;

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

    private sealed record SettingsShortcutItem(string Key, string Description);

    private static Brush? TryGetBrush(string resourceKey)
    {
        return Application.Current.Resources.TryGetValue(resourceKey, out var value) ? value as Brush : null;
    }

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
        SetMenuTitle(FileMenuBarItem, T("Fichier", "File"));
        SetMenuTitle(EditMenuBarItem, T("Édition", "Edit"));
        SetMenuTitle(ImageMenuBarItem, T("Image", "Image"));
        SetMenuTitle(ExportMenuBarItem, T("Export", "Export"));
        SetMenuTitle(HelpMenuBarItem, T("Aide", "Help"));
        SetMenuTooltip(FileMenuBarItem, T("Ouvrir les actions de fichier", "Open file actions"));
        SetMenuTooltip(EditMenuBarItem, T("Ouvrir les actions d’édition", "Open edit actions"));
        SetMenuTooltip(ImageMenuBarItem, T("Ouvrir les actions image", "Open image actions"));
        SetMenuTooltip(ExportMenuBarItem, T("Ouvrir les options d’export", "Open export options"));
        SetMenuTooltip(HelpMenuBarItem, T("Ouvrir l’aide et les raccourcis", "Open help and shortcuts"));

        SetMenuText(NewProjectMenuItem, T("Nouveau", "New"));
        SetMenuText(OpenProjectMenuItem, T("Ouvrir", "Open"));
        SetMenuText(SaveProjectMenuItem, T("Sauver", "Save"));
        SetMenuText(UndoMenuItem, T("Annuler", "Undo"));
        SetMenuText(RedoMenuItem, T("Refaire", "Redo"));
        SetMenuText(ImportScreenshotMenuItem, T("Importer", "Import"));
        SetMenuText(OcrMenuItem, "OCR");
        SetMenuText(ExportCurrentMenuItem, T("PNG de l’étape courante", "Current step PNG"));
        SetMenuText(ExportAllFormatsMenuItem, T("Tout exporter", "Export all"));
        SetMenuText(ExportMarkdownMenuItem, "Markdown");
        SetMenuText(ExportHtmlMenuItem, "HTML");
        SetMenuText(ExportPdfMenuItem, "PDF");
        SetMenuText(ExportDocxMenuItem, T("Word", "Word"));
        SetMenuText(PublishNotionMenuItem, T("Vers Notion", "To Notion"));
        SetMenuText(PublishConfluenceMenuItem, T("Vers Confluence", "To Confluence"));
        SetMenuText(PublishSharePointMenuItem, T("Vers SharePoint", "To SharePoint"));
        SetMenuText(OpenHelpMenuItem, T("Aide rapide", "Quick help"));
        SetMenuText(PrintScreenMenuItem, T("Impr écran", "Print Screen"));
        SetMenuText(AltPrintScreenMenuItem, T("Alt + Impr écran", "Alt + Print Screen"));
        SetMenuText(WinPrintScreenMenuItem, T("Win + Impr écran", "Win + Print Screen"));
        SetMenuText(WinShiftSMenuItem, T("Win + Shift + S", "Win + Shift + S"));
        SetMenuText(FnWinSpaceMenuItem, T("Fn + Win + Espace", "Fn + Win + Space"));
        SetButtonContent(SettingsMenuButton, T("⚙ Réglages", "⚙ Settings"));
        SetButtonTooltip(SettingsMenuButton, T("Ouvrir les réglages", "Open settings"));
        SetButtonTooltip(AppMenuButton, T("Menu de l’application", "Application menu"));
        SetButtonTooltip(DocumentTitleEditButton, T("Modifier le titre du document", "Edit the document title"));
        SetButtonContent(InspectorToggleButton, T("◫", "◫"));
        SetButtonTooltip(InspectorToggleButton, T("Afficher ou masquer l’inspecteur", "Show or hide the inspector"));

        SetText(SettingsGeneralNavText, T("Général", "General"));
        SetText(SettingsCaptureNavText, T("Capture", "Capture"));
        SetText(SettingsExportNavText, T("Export", "Export"));
        SetText(SettingsPublicationNavText, T("Publication", "Publication"));
        SetText(SettingsShortcutsNavText, T("Raccourcis", "Shortcuts"));
        SetText(SettingsAboutNavText, T("À propos", "About"));

        SetCommandButton(NewProjectButton, T("Nouveau", "New"), T("Nouveau projet", "New project"));
        SetCommandButton(OpenProjectButton, T("Ouvrir", "Open"), T("Ouvrir un projet", "Open a project"));
        SetCommandButton(SaveProjectButton, T("Sauver", "Save"), T("Sauvegarder le projet", "Save project"));
        SetCommandButton(DemoProjectButton, T("Demo", "Demo"), T("Charger le modèle de démonstration", "Load the demo template"));
        SetButtonContent(CaptureQuickImportButton, T("Importer", "Import"));
        SetButtonContent(ImportScreenshotButton, T("Importer", "Import"));
        SetCommandButton(OcrButton, T("OCR", "OCR"), T("Reconnaître le texte de la capture courante", "Recognize text from the current image"));
        SetButtonContent(ExportCurrentButton, T("Exporter", "Export"));
        SetCommandButton(UndoButton, T("Annuler", "Undo"), T("Annuler la dernière action", "Undo the last action"));
        SetCommandButton(RedoButton, T("Rétablir", "Redo"), T("Rétablir la dernière action", "Redo the last action"));
        SetButtonContent(CollapseStepsButton, T("Cacher étapes", "Hide steps"));
        SetButtonContent(BrowseExportFolderButton, T("Parcourir", "Browse"));
        SetButtonContent(BrowseSharePointFolderButton, T("Parcourir", "Browse"));
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
        SetCommandButton(ApplyCropButton, T("Appliquer", "Apply"), T("Appliquer le crop", "Apply crop"));
        SetCommandButton(ResetCropButton, T("Réinitialiser", "Reset"), T("Réinitialiser le crop", "Reset crop"));
        SetCommandButton(ClearAnnotationsButton, T("Effacer tout", "Clear all"), T("Effacer toutes les annotations", "Clear all annotations"));
        SetCommandButton(DuplicateSelectedButton, T("Dupliquer", "Duplicate"), T("Dupliquer l’objet sélectionné", "Duplicate selected object"));
        SetCommandButton(DeleteSelectedButton, T("Supprimer", "Delete"), T("Supprimer l’objet sélectionné", "Delete selected object"));
        SetCommandButton(BringForwardButton, T("Devant", "Front"), T("Envoyer l’objet devant", "Bring the object to front"));
        SetCommandButton(SendBackwardButton, T("Derrière", "Back"), T("Envoyer l’objet derrière", "Send the object to back"));
        SetCommandButton(AddLegendButton, T("+ Légende", "+ Legend"), T("Ajouter une légende", "Add a legend"));
        SetCommandButton(PublishNotionButton, T("Publier vers Notion", "Publish to Notion"), T("Publier le document courant vers Notion", "Publish the current guide to Notion"));
        SetCommandButton(PublishConfluenceButton, T("Publier vers Confluence", "Publish to Confluence"), T("Publier le document courant vers Confluence", "Publish the current guide to Confluence"));
        SetCommandButton(PublishSharePointButton, T("Publier vers SharePoint", "Publish to SharePoint"), T("Copier le package vers le dossier synchronisé", "Copy the package to the synchronized folder"));
        SetButtonContent(AddStepButton, T("Ajouter", "Add"));
        SetButtonContent(CollapseStepsButton, T("‹", "‹"));
        SetButtonTooltip(CollapseStepsButton, T("Masquer les étapes", "Hide the steps"));

        SetCommandButton(ToolbarZoomOutButton, T("−", "−"), T("Rétrécir la vue", "Zoom out"));
        SetCommandButton(ToolbarZoomResetButton, T("100 %", "100 %"), T("Remettre à 100 %", "Reset to 100%"));
        SetCommandButton(ToolbarZoomInButton, T("+", "+"), T("Agrandir la vue", "Zoom in"));
        SetCommandButton(ToolbarFitButton, T("Ajuster", "Fit"), T("Ajuster la vue", "Fit to view"));
        SetCommandButton(ToolbarStepsButton, T("Étapes", "Steps"), T("Afficher ou masquer les étapes", "Show or hide the steps pane"));

        SetToolButton(ToolbarSelectButton, EditorTool.Select, T("Sélection", "Select"));
        SetToolButton(ToolbarCropButton, EditorTool.Crop, T("Recadrer la capture", "Crop the capture"));
        SetOcrButton(ToolbarOcrButton, T("OCR", "OCR"));
        SetToolButton(ToolbarTextButton, EditorTool.Text, T("Texte", "Text"));
        SetToolButton(ToolbarStickerButton, EditorTool.Sticker, T("Gommette", "Sticker"));
        SetToolButton(ToolbarArrowStraightButton, EditorTool.ArrowStraight, T("Flèche droite", "Straight arrow"));
        SetToolButton(ToolbarArrowCurvedButton, EditorTool.ArrowCurved, T("Flèche courbe", "Curved arrow"));
        SetToolButton(ToolbarRectangleButton, EditorTool.Rectangle, T("Rectangle", "Rectangle"));
        SetToolButton(ToolbarEllipseButton, EditorTool.Ellipse, T("Ovale", "Ellipse"));
        SetToolButton(ToolbarFocusButton, EditorTool.Focus, T("Focus", "Focus"));
        SetToolButton(ToolbarMaskButton, EditorTool.Mask, T("Masquage", "Mask"));
        SetToolButton(ToolbarToolPickerButton, _currentTool, T("Choisir un outil", "Choose a tool"));

        SetText(CapturePageTitleText, T("Aide", "Help"));
        SetText(CapturePageSummaryText, T("Raccourcis de capture Windows et raccourcis SnapSlate utiles.", "Windows capture shortcuts and useful SnapSlate shortcuts."));
        SetText(CaptureHowTitleText, T("Comment ça marche", "How it works"));
        SetText(CaptureHowBodyText, T("Impr écran, Alt + Impr écran, Win + Impr écran, Win + Shift + S ou Fn + Win + Espace : toutes les façons courantes de capturer avant import.", "Print Screen, Alt + Print Screen, Win + Print Screen, Win + Shift + S, or Fn + Win + Space: the common ways to capture before import."));
        SetText(CaptureStepOneTitleText, T("Impr écran", "Print Screen"));
        SetText(CaptureStepOneBodyText, T("Copie l’écran entier dans le presse-papiers.", "Copies the whole screen to the clipboard."));
        SetText(CaptureStepTwoTitleText, T("Alt + Impr écran", "Alt + Print Screen"));
        SetText(CaptureStepTwoBodyText, T("Copie uniquement la fenêtre active.", "Copies only the active window."));
        SetText(CaptureStepThreeTitleText, T("Win + Impr écran", "Win + Print Screen"));
        SetText(CaptureStepThreeBodyText, T("Sauvegarde la capture dans Images > Captures d’écran.", "Saves the capture to Pictures > Screenshots."));
        SetText(CaptureShortcutsTitleText, T("Raccourcis de capture", "Capture shortcuts"));
        SetText(CaptureShortcutOneText, T("Win + Shift + S : outil Capture d’écran avec modes forme libre, rectangle, fenêtre et plein écran.", "Win + Shift + S: Snipping Tool with freeform, rectangle, window, and full-screen modes."));
        SetText(CaptureShortcutTwoText, T("Fn + Win + Espace : utile sur les claviers sans touche Impr écran.", "Fn + Win + Space: useful on keyboards without a Print Screen key."));
        SetText(CaptureShortcutThreeText, T("PrtScn, Alt + PrtScn et Win + PrtScn fonctionnent selon le matériel.", "PrtScn, Alt + PrtScn, and Win + PrtScn work depending on the hardware."));
        SetText(CaptureShortcutFourText, T("Les captures importées peuvent aussi arriver automatiquement depuis le presse-papiers.", "Imported screenshots can also arrive automatically from the clipboard."));
        SetText(CaptureShortcutFiveText, T("Ctrl + O : importer un fichier image.", "Ctrl + O: import an image file."));
        SetText(HelpShortcutsTitleText, T("Raccourcis utiles", "Useful shortcuts"));
        SetText(CaptureStatusTitleText, T("État", "Status"));
        SetText(CaptureStatusBodyText, T("L’import automatique fonctionne tant que SnapSlate est ouvert.", "Automatic import works while SnapSlate is open."));
        SetText(ProcedureTitleLabelText, T("Document", "Document"));
        SetText(ToolbarDocumentLabelText, T("Document", "Document"));
        SetText(CurrentToolLabelText, T("Outil", "Tool"));
        SetText(ToolbarCurrentToolLabelText, T("Outil :", "Tool:"));
        SetText(StepsPaneHintText, T("Glissez pour réordonner les étapes.", "Drag to reorder the steps."));
        SetText(PaletteSectionTitleText, T("Palette", "Palette"));
        SetText(ToolbarPaletteLabelText, T("Palette", "Palette"));
        SetText(ShadeSectionTitleText, T("Nuances", "Shades"));
        SetText(ToolbarOpacityLabelText, T("Opacité", "Opacity"));
        SetText(ToolbarOpacityValueText, "100 %");
        SetText(StepsPaneTitleText, T("Étapes", "Steps"));
        SetText(StepDetailsSubtitleText, T("Titre et note de l’étape", "Step title and note"));
        SetText(SelectedColorLabelText, T("Couleur", "Color"));
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
        SetText(PublishPageTitleText, T("Publication", "Publish"));
        SetText(PublishPageSummaryText, T("OCR depuis la barre d’outils, puis publication vers Notion, Confluence ou un dossier SharePoint synchronisé.", "Use OCR from the toolbar, then publish to Notion, Confluence, or a synchronized SharePoint folder."));
        SetText(OcrSectionTitleText, T("OCR", "OCR"));
        SetText(OcrSectionHelpText, T("Le texte reconnu est ajouté comme annotation texte sur la capture courante.", "Recognized text is inserted as a text annotation on the current capture."));
        SetText(NotionSectionTitleText, T("Notion", "Notion"));
        SetText(NotionSectionHelpText, T("Partagez la page parente avec votre intégration Notion, puis saisissez l’ID de la page parente et le jeton API.", "Share the parent page with your Notion integration, then enter the parent page ID and API token."));
        SetText(NotionParentPageLabelText, T("ID de page parente", "Parent page ID"));
        SetText(NotionTokenLabelText, T("Jeton API Notion", "Notion API token"));
        SetText(ConfluenceSectionTitleText, T("Confluence", "Confluence"));
        SetText(ConfluenceSectionHelpText, T("Créez une page dans un espace Confluence Cloud avec l’URL du site, la clé d’espace, l’e-mail et le jeton API.", "Create a page in a Confluence Cloud space using the site URL, space key, email, and API token."));
        SetText(ConfluenceBaseUrlLabelText, T("URL du site", "Site URL"));
        SetText(ConfluenceSpaceKeyLabelText, T("Clé d’espace", "Space key"));
        SetText(ConfluenceParentPageLabelText, T("ID de page parente (optionnel)", "Parent page ID (optional)"));
        SetText(ConfluenceEmailLabelText, T("E-mail du compte", "Account email"));
        SetText(ConfluenceTokenLabelText, T("Jeton API Confluence", "Confluence API token"));
        SetText(SharePointSectionTitleText, T("SharePoint", "SharePoint"));
        SetText(SharePointSectionHelpText, T("Copiez le package d’export dans un dossier synchronisé avec SharePoint ou OneDrive.", "Copy the export package into a folder synchronized with SharePoint or OneDrive."));
        SetText(SharePointFolderLabelText, T("Dossier SharePoint synchronisé", "Synchronized SharePoint folder"));
        SetText(SettingsPageTitleText, T("Général", "General"));
        SetText(AppearanceSettingsTitleText, T("Apparence", "Appearance"));
        SetText(LanguageSettingsTitleText, T("Langue", "Language"));
        SetText(CaptureSettingsTitleText, T("Capture", "Capture"));
        SetText(ThemeSettingLabelText, T("Thème", "Theme"));
        SetText(LanguageSettingLabelText, T("Langue de l’application", "App language"));
        SetText(CaptureSettingLabelText, T("Import automatique", "Automatic import"));
        SetText(ThemeSettingHelpText, T("Le mode Système suit le thème Windows automatiquement.", "System follows Windows automatically."));
        SetText(LanguageSettingHelpText, T("Par défaut, SnapSlate suit la langue de Windows.", "By default, SnapSlate follows the Windows language."));
        SetText(CaptureSettingHelpText, T("Active l’import automatique des captures copiées dans le presse-papiers.", "Enables automatic import of captures copied to the clipboard."));
        SetText(SettingsRestartHintText, T("Le thème et la langue s’appliquent immédiatement, sans redémarrage.", "Theme and language apply immediately without restarting."));
        SetText(SettingsAboutTitleText, T("À propos", "About"));
        SetText(SettingsAboutBodyText, T("SnapSlate accélère la création de procédures visuelles avec captures, annotations, export et publication.", "SnapSlate accelerates visual procedure creation with captures, annotations, export, and publishing."));
        SetButtonContent(SettingsCancelButton, T("Annuler", "Cancel"));
        SetButtonContent(SettingsSaveButton, T("Enregistrer", "Save"));
        SetText(SettingsTipTitleText, T("Conseil SnapSlate", "SnapSlate tip"));
        SetText(SettingsTipBodyText, T("Les raccourcis utiles sont regroupés dans Aide pour rester visibles sans encombrer les réglages.", "Useful shortcuts are grouped in Help so they stay visible without crowding Settings."));
        AutoImportClipboardToggleSwitch.OnContent = string.Empty;
        AutoImportClipboardToggleSwitch.OffContent = string.Empty;
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
        HelpShortcutsListView.ItemsSource = new[]
        {
            new SettingsShortcutItem("Ctrl + Z", T("Annuler la dernière action", "Undo the last action")),
            new SettingsShortcutItem("Ctrl + Y", T("Rétablir la dernière action", "Redo the last action")),
            new SettingsShortcutItem("Ctrl + S", T("Sauvegarder le projet", "Save the project")),
            new SettingsShortcutItem("Ctrl + O", T("Ouvrir un projet", "Open a project")),
            new SettingsShortcutItem("Ctrl + N", T("Nouveau projet", "New project")),
            new SettingsShortcutItem("Ctrl + Shift + N", T("Modèle de démonstration", "Demo template")),
            new SettingsShortcutItem("Win + Shift + S", T("Importer la capture active", "Import the current capture")),
            new SettingsShortcutItem("Suppr", T("Supprimer l’objet sélectionné", "Delete the selected object"))
        };
    }

    private string GetThemeDisplayName(ThemeChoice theme)
    {
        return theme switch
        {
            ThemeChoice.Light => T("Clair", "Light"),
            ThemeChoice.Dark => T("Sombre", "Dark"),
            _ => T("Système", "System")
        };
    }

    private string GetLanguageDisplayName(LanguageChoice language)
    {
        return language switch
        {
            LanguageChoice.French => T("Français", "French"),
            LanguageChoice.English => T("English", "English"),
            _ => T("Suivre Windows", "Follow Windows")
        };
    }

    private string GetExportFormatDisplayName(ExportFormatChoice format)
    {
        return format switch
        {
            ExportFormatChoice.Html => "HTML",
            ExportFormatChoice.Markdown => "Markdown",
            ExportFormatChoice.Pdf => "PDF",
            ExportFormatChoice.Docx => T("Word", "Word"),
            _ => "PNG"
        };
    }

    private void UpdateSettingsSummary()
    {
        // The old settings summary pane has been removed.
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
        ExportFormatComboBox.SelectedIndex = (int)App.Preferences.DefaultExportFormat;
        AutoImportClipboardToggleSwitch.IsOn = App.Preferences.AutoImportClipboard;
        ExportFolderTextBox.Text = App.Preferences.ExportFolder;
        NotionParentPageIdTextBox.Text = App.Preferences.NotionParentPageId;
        ConfluenceBaseUrlTextBox.Text = App.Preferences.ConfluenceBaseUrl;
        ConfluenceSpaceKeyTextBox.Text = App.Preferences.ConfluenceSpaceKey;
        ConfluenceParentPageIdTextBox.Text = App.Preferences.ConfluenceParentPageId;
        ConfluenceEmailTextBox.Text = App.Preferences.ConfluenceEmail;
        SharePointSyncFolderTextBox.Text = App.Preferences.SharePointSyncFolder;
        ApplyProcedurePaneLayout();
        ApplyFloatingToolRailPosition(App.Preferences.FloatingToolRailLeft, App.Preferences.FloatingToolRailTop);
        UpdateSettingsSummary();
    }

    private void ApplyProcedurePaneLayout()
    {
        var stepsCollapsed = App.Preferences.CollapseStepsPane;
        StepsPaneColumn.Width = new GridLength(stepsCollapsed ? 48 : 256);
        InspectorPaneColumn.Width = new GridLength(_isInspectorPaneCollapsed ? 0 : 312);
        StepsHeaderPanel.Visibility = stepsCollapsed ? Visibility.Collapsed : Visibility.Visible;
        DocumentTabsListView.Visibility = stepsCollapsed ? Visibility.Collapsed : Visibility.Visible;
        DocumentTabsCompactListView.Visibility = stepsCollapsed ? Visibility.Visible : Visibility.Collapsed;
        CollapseStepsButton.Visibility = stepsCollapsed ? Visibility.Collapsed : Visibility.Visible;
        AddStepButton.Visibility = stepsCollapsed ? Visibility.Collapsed : Visibility.Visible;
        StepsPaneHintText.Visibility = stepsCollapsed ? Visibility.Collapsed : Visibility.Visible;
        InspectorToggleButton.Content = _isInspectorPaneCollapsed ? T("◫", "◫") : T("◫", "◫");
        ToolbarStepsButton.Content = stepsCollapsed ? T("Étapes", "Steps") : T("Étapes", "Steps");
        ToolbarOpacityValueText.Text = string.Format(CultureInfo.CurrentCulture, "{0:0} %", ToolbarOpacitySlider.Value * 100);
    }

    private void ApplyFloatingToolRailPosition(double left, double top)
    {
        var railWidth = FloatingToolRail.ActualWidth > 0 ? FloatingToolRail.ActualWidth : 72;
        var railHeight = FloatingToolRail.ActualHeight > 0 ? FloatingToolRail.ActualHeight : 420;
        var viewportWidth = PreviewViewport.ActualWidth > 0 ? PreviewViewport.ActualWidth : PreviewViewport.Width;
        var viewportHeight = PreviewViewport.ActualHeight > 0 ? PreviewViewport.ActualHeight : PreviewViewport.Height;
        var maxLeft = Math.Max(8, viewportWidth - railWidth - 8);
        var maxTop = Math.Max(8, viewportHeight - railHeight - 8);

        var clampedLeft = Math.Clamp(left, 8, maxLeft);
        var clampedTop = Math.Clamp(top, 8, maxTop);

        FloatingToolRail.Margin = new Thickness(clampedLeft, clampedTop, 0, 0);
        App.Preferences.FloatingToolRailLeft = clampedLeft;
        App.Preferences.FloatingToolRailTop = clampedTop;
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

    private void SetCommandButton(Button button, string content, string tooltip)
    {
        button.Content = content;
        ToolTipService.SetToolTip(button, tooltip);
        AutomationProperties.SetName(button, tooltip);
    }

    private void SetButtonTooltip(Button button, string tooltip)
    {
        ToolTipService.SetToolTip(button, tooltip);
        AutomationProperties.SetName(button, tooltip);
    }

    private void SetToolButton(Button button, EditorTool tool, string tooltip)
    {
        _iconBrushOverride = CreateToolAccentBrush(tool);
        button.Content = CreateToolIcon(tool);
        _iconBrushOverride = null;
        button.MinWidth = 38;
        button.MinHeight = 38;
        button.Padding = new Thickness(0);
        button.Background = GetBrush("ShellPanelBrush");
        button.BorderBrush = GetBrush("ShellStrokeBrush");
        button.BorderThickness = new Thickness(1);
        button.CornerRadius = new CornerRadius(10);
        ToolTipService.SetToolTip(button, tooltip);
        AutomationProperties.SetName(button, tooltip);
    }

    private void SetOcrButton(Button button, string tooltip)
    {
        _iconBrushOverride = CreateOcrAccentBrush();
        button.Content = CreateOcrIcon();
        _iconBrushOverride = null;
        button.MinWidth = 38;
        button.MinHeight = 38;
        button.Padding = new Thickness(0);
        button.Background = GetBrush("ShellPanelBrush");
        button.BorderBrush = GetBrush("ShellStrokeBrush");
        button.BorderThickness = new Thickness(1);
        button.CornerRadius = new CornerRadius(10);
        ToolTipService.SetToolTip(button, tooltip);
        AutomationProperties.SetName(button, tooltip);
    }

    private static Brush CreateToolAccentBrush(EditorTool tool)
    {
        var color = tool switch
        {
            EditorTool.Select => ColorHelper.FromArgb(255, 84, 98, 122),
            EditorTool.Crop => ColorHelper.FromArgb(255, 38, 123, 245),
            EditorTool.Text => ColorHelper.FromArgb(255, 20, 176, 185),
            EditorTool.Sticker => ColorHelper.FromArgb(255, 232, 64, 149),
            EditorTool.ArrowStraight => ColorHelper.FromArgb(255, 255, 159, 55),
            EditorTool.ArrowCurved => ColorHelper.FromArgb(255, 141, 92, 242),
            EditorTool.Rectangle => ColorHelper.FromArgb(255, 47, 132, 241),
            EditorTool.Ellipse => ColorHelper.FromArgb(255, 25, 175, 190),
            EditorTool.Focus => ColorHelper.FromArgb(255, 48, 81, 214),
            EditorTool.Mask => ColorHelper.FromArgb(255, 239, 83, 80),
            _ => ColorHelper.FromArgb(255, 112, 118, 132)
        };

        return new SolidColorBrush(color);
    }

    private static Brush CreateOcrAccentBrush()
    {
        return new SolidColorBrush(ColorHelper.FromArgb(255, 118, 72, 231));
    }

    private static Brush CreateToolBrush(EditorTool tool)
    {
        var (start, end) = tool switch
        {
            EditorTool.Select => (ColorHelper.FromArgb(255, 84, 98, 122), ColorHelper.FromArgb(255, 54, 71, 97)),
            EditorTool.Crop => (ColorHelper.FromArgb(255, 40, 145, 255), ColorHelper.FromArgb(255, 16, 98, 219)),
            EditorTool.Text => (ColorHelper.FromArgb(255, 33, 201, 208), ColorHelper.FromArgb(255, 17, 166, 182)),
            EditorTool.Sticker => (ColorHelper.FromArgb(255, 255, 93, 170), ColorHelper.FromArgb(255, 235, 55, 136)),
            EditorTool.ArrowStraight => (ColorHelper.FromArgb(255, 255, 194, 77), ColorHelper.FromArgb(255, 255, 153, 35)),
            EditorTool.ArrowCurved => (ColorHelper.FromArgb(255, 160, 103, 255), ColorHelper.FromArgb(255, 120, 63, 233)),
            EditorTool.Rectangle => (ColorHelper.FromArgb(255, 62, 157, 255), ColorHelper.FromArgb(255, 30, 110, 234)),
            EditorTool.Ellipse => (ColorHelper.FromArgb(255, 40, 188, 205), ColorHelper.FromArgb(255, 20, 158, 174)),
            EditorTool.Focus => (ColorHelper.FromArgb(255, 58, 97, 224), ColorHelper.FromArgb(255, 33, 56, 196)),
            EditorTool.Mask => (ColorHelper.FromArgb(255, 255, 97, 97), ColorHelper.FromArgb(255, 232, 45, 45)),
            _ => (ColorHelper.FromArgb(255, 120, 120, 120), ColorHelper.FromArgb(255, 80, 80, 80))
        };

        var brush = new LinearGradientBrush
        {
            StartPoint = new Point(0, 0),
            EndPoint = new Point(1, 1)
        };
        brush.GradientStops.Add(new GradientStop { Color = start, Offset = 0 });
        brush.GradientStops.Add(new GradientStop { Color = end, Offset = 1 });
        return brush;
    }

    private static Brush CreateOcrBrush()
    {
        var brush = new LinearGradientBrush
        {
            StartPoint = new Point(0, 0),
            EndPoint = new Point(1, 1)
        };
        brush.GradientStops.Add(new GradientStop { Color = ColorHelper.FromArgb(255, 149, 96, 255), Offset = 0 });
        brush.GradientStops.Add(new GradientStop { Color = ColorHelper.FromArgb(255, 106, 50, 229), Offset = 1 });
        return brush;
    }

    private static FrameworkElement CreateToolIcon(EditorTool tool)
    {
        return tool switch
        {
            EditorTool.Select => CreateSelectIcon(),
            EditorTool.Crop => CreateCropIcon(),
            EditorTool.Text => CreateTextIcon(),
            EditorTool.Sticker => CreateStickerIcon(),
            EditorTool.ArrowStraight => CreateStraightArrowIcon(),
            EditorTool.ArrowCurved => CreateCurvedArrowIcon(),
            EditorTool.Rectangle => CreateRectangleIcon(),
            EditorTool.Ellipse => CreateEllipseIcon(),
            EditorTool.Focus => CreateFocusIcon(),
            EditorTool.Mask => CreateMaskIcon(),
            _ => CreateFallbackIcon()
        };
    }

    private static FrameworkElement CreateSelectIcon()
    {
        var geometry = new PathGeometry();
        var figure = new PathFigure
        {
            StartPoint = new Point(3, 2),
            IsClosed = true,
            IsFilled = true
        };

        figure.Segments.Add(new LineSegment { Point = new Point(3, 16) });
        figure.Segments.Add(new LineSegment { Point = new Point(7, 12) });
        figure.Segments.Add(new LineSegment { Point = new Point(10, 18) });
        figure.Segments.Add(new LineSegment { Point = new Point(12, 17) });
        figure.Segments.Add(new LineSegment { Point = new Point(9, 11) });
        figure.Segments.Add(new LineSegment { Point = new Point(15, 11) });
        geometry.Figures.Add(figure);

        return new XamlPath
        {
            Data = geometry,
            Fill = CreateIconBrush(),
            Stroke = CreateIconBrush(),
            StrokeThickness = 0.75,
            StrokeLineJoin = PenLineJoin.Round,
            StrokeStartLineCap = PenLineCap.Round,
            StrokeEndLineCap = PenLineCap.Round,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
    }

    private static Canvas CreateIconCanvas()
    {
        return new Canvas
        {
            Width = 20,
            Height = 20,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
    }

    private static Brush CreateIconBrush() => _iconBrushOverride ?? new SolidColorBrush(Colors.White);

    private static void AddLine(Canvas canvas, double x1, double y1, double x2, double y2, double thickness = 2)
    {
        canvas.Children.Add(new XamlLine
        {
            X1 = x1,
            Y1 = y1,
            X2 = x2,
            Y2 = y2,
            Stroke = CreateIconBrush(),
            StrokeThickness = thickness,
            StrokeStartLineCap = PenLineCap.Round,
            StrokeEndLineCap = PenLineCap.Round
        });
    }

    private static FrameworkElement CreateCropIcon()
    {
        var canvas = CreateIconCanvas();
        AddLine(canvas, 3, 7, 8, 7);
        AddLine(canvas, 7, 3, 7, 8);
        AddLine(canvas, 12, 3, 12, 8);
        AddLine(canvas, 12, 7, 17, 7);
        AddLine(canvas, 3, 13, 8, 13);
        AddLine(canvas, 7, 12, 7, 17);
        AddLine(canvas, 12, 12, 12, 17);
        AddLine(canvas, 12, 13, 17, 13);
        return canvas;
    }

    private static FrameworkElement CreateOcrIcon()
    {
        var canvas = CreateIconCanvas();

        var document = new XamlRectangle
        {
            Width = 10.8,
            Height = 13.8,
            RadiusX = 2,
            RadiusY = 2,
            Stroke = CreateIconBrush(),
            StrokeThickness = 1.8,
            Fill = new SolidColorBrush(Color.FromArgb(28, 255, 255, 255))
        };
        Canvas.SetLeft(document, 4.6);
        Canvas.SetTop(document, 3.1);

        var foldedCorner = new XamlPath
        {
            Data = new PathGeometry
            {
                Figures =
                {
                    new PathFigure
                    {
                        StartPoint = new Point(12, 4),
                        IsClosed = false,
                        Segments =
                        {
                            new LineSegment { Point = new Point(16, 8) },
                            new LineSegment { Point = new Point(12, 8) }
                        }
                    }
                }
            },
            Stroke = CreateIconBrush(),
            StrokeThickness = 1.5,
            StrokeLineJoin = PenLineJoin.Round,
            StrokeStartLineCap = PenLineCap.Round,
            StrokeEndLineCap = PenLineCap.Round,
            Fill = new SolidColorBrush(Colors.Transparent)
        };

        var label = new TextBlock
        {
            Text = "OCR",
            Foreground = CreateIconBrush(),
            FontFamily = new FontFamily("Segoe UI Semibold"),
            FontSize = 5.8,
            Width = 10.5,
            Height = 5.5
        };
        Canvas.SetLeft(label, 4.7);
        Canvas.SetTop(label, 8.0);

        AddLine(canvas, 2.5, 5, 5, 5, 1.6);
        AddLine(canvas, 5, 2.5, 5, 5, 1.6);
        AddLine(canvas, 15, 5, 17.5, 5, 1.6);
        AddLine(canvas, 15, 2.5, 15, 5, 1.6);
        AddLine(canvas, 2.5, 15, 5, 15, 1.6);
        AddLine(canvas, 5, 15, 5, 17.5, 1.6);
        AddLine(canvas, 15, 15, 17.5, 15, 1.6);
        AddLine(canvas, 15, 15, 15, 17.5, 1.6);

        canvas.Children.Add(document);
        canvas.Children.Add(foldedCorner);
        canvas.Children.Add(label);
        return canvas;
    }

    private static FrameworkElement CreateTextIcon()
    {
        var canvas = CreateIconCanvas();
        var text = new TextBlock
        {
            Text = "T",
            Foreground = CreateIconBrush(),
            FontFamily = new FontFamily("Segoe UI Semibold"),
            FontSize = 17,
            Width = 9,
            Height = 16
        };
        Canvas.SetLeft(text, 4.8);
        Canvas.SetTop(text, 0.8);
        canvas.Children.Add(text);
        AddLine(canvas, 15, 4, 15, 16, 1.8);
        canvas.Children.Add(new XamlLine
        {
            X1 = 13.2,
            Y1 = 4,
            X2 = 13.2,
            Y2 = 16,
            Stroke = CreateIconBrush(),
            StrokeThickness = 1.8,
            StrokeDashArray = new DoubleCollection { 1, 1.6 },
            StrokeStartLineCap = PenLineCap.Round,
            StrokeEndLineCap = PenLineCap.Round
        });
        AddLine(canvas, 5, 16, 13.5, 16, 1.8);
        return canvas;
    }

    private static FrameworkElement CreateStickerIcon()
    {
        var grid = new Grid
        {
            Width = 20,
            Height = 20,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        var bubble = new XamlEllipse
        {
            Width = 12,
            Height = 12,
            Stroke = CreateIconBrush(),
            StrokeThickness = 1.8,
            Fill = new SolidColorBrush(Color.FromArgb(32, 255, 255, 255)),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Top,
            Margin = new Thickness(0, 2, 0, 0)
        };

        var tail = new XamlPolygon
        {
            Points = new PointCollection
            {
                new Point(10, 18),
                new Point(7.5, 13.5),
                new Point(12.5, 13.5)
            },
            Stroke = CreateIconBrush(),
            StrokeThickness = 1.6,
            Fill = new SolidColorBrush(Colors.Transparent),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        var number = new TextBlock
        {
            Text = "1",
            Foreground = CreateIconBrush(),
            FontFamily = new FontFamily("Segoe UI Semibold"),
            FontSize = 10,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, -1, 0, 0)
        };

        grid.Children.Add(bubble);
        grid.Children.Add(tail);
        grid.Children.Add(number);
        return grid;
    }

    private static FrameworkElement CreateStraightArrowIcon()
    {
        var canvas = CreateIconCanvas();
        AddLine(canvas, 3, 10, 15, 10, 2.2);
        AddLine(canvas, 12, 6.5, 16, 10, 2.2);
        AddLine(canvas, 12, 13.5, 16, 10, 2.2);
        return canvas;
    }

    private static FrameworkElement CreateCurvedArrowIcon()
    {
        var canvas = CreateIconCanvas();
        var path = new XamlPath
        {
            Stroke = CreateIconBrush(),
            StrokeThickness = 2.2,
            StrokeLineJoin = PenLineJoin.Round,
            StrokeStartLineCap = PenLineCap.Round,
            StrokeEndLineCap = PenLineCap.Round,
            Data = new PathGeometry
            {
                Figures =
                {
                    new PathFigure
                    {
                        StartPoint = new Point(4, 15),
                        IsClosed = false,
                        IsFilled = false,
                        Segments =
                        {
                            new BezierSegment
                            {
                                Point1 = new Point(4, 8),
                                Point2 = new Point(9.5, 4.5),
                                Point3 = new Point(15.5, 6.5)
                            }
                        }
                    }
                }
            }
        };

        canvas.Children.Add(path);
        AddLine(canvas, 12, 4.8, 15.5, 6.5, 2.2);
        AddLine(canvas, 13.2, 9.2, 15.5, 6.5, 2.2);
        return canvas;
    }

    private static FrameworkElement CreateRectangleIcon()
    {
        return new XamlRectangle
        {
            Width = 12,
            Height = 9,
            RadiusX = 1.4,
            RadiusY = 1.4,
            Stroke = CreateIconBrush(),
            StrokeThickness = 2.1,
            Fill = new SolidColorBrush(Colors.Transparent),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
    }

    private static FrameworkElement CreateEllipseIcon()
    {
        return new XamlEllipse
        {
            Width = 12,
            Height = 9,
            Stroke = CreateIconBrush(),
            StrokeThickness = 2.1,
            Fill = new SolidColorBrush(Colors.Transparent),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
    }

    private static FrameworkElement CreateFocusIcon()
    {
        var canvas = CreateIconCanvas();
        AddLine(canvas, 2.5, 5, 5.5, 5, 1.9);
        AddLine(canvas, 5, 2.5, 5, 5.5, 1.9);
        AddLine(canvas, 14.5, 2.5, 14.5, 5.5, 1.9);
        AddLine(canvas, 14.5, 5, 17.5, 5, 1.9);
        AddLine(canvas, 2.5, 15, 5.5, 15, 1.9);
        AddLine(canvas, 5, 15, 5, 17.5, 1.9);
        AddLine(canvas, 14.5, 15, 14.5, 17.5, 1.9);
        AddLine(canvas, 14.5, 15, 17.5, 15, 1.9);

        var focusCircle = new XamlEllipse
        {
            Width = 6,
            Height = 6,
            Stroke = CreateIconBrush(),
            StrokeThickness = 2,
            Fill = new SolidColorBrush(Color.FromArgb(30, 255, 255, 255)),
        };
        Canvas.SetLeft(focusCircle, 7);
        Canvas.SetTop(focusCircle, 7);
        canvas.Children.Add(focusCircle);

        return canvas;
    }

    private static FrameworkElement CreateMaskIcon()
    {
        var canvas = CreateIconCanvas();
        var document = new XamlRectangle
        {
            Width = 12,
            Height = 14,
            RadiusX = 1.5,
            RadiusY = 1.5,
            Stroke = CreateIconBrush(),
            StrokeThickness = 1.8,
            Fill = new SolidColorBrush(Color.FromArgb(24, 255, 255, 255)),
        };
        Canvas.SetLeft(document, 4);
        Canvas.SetTop(document, 3);
        canvas.Children.Add(document);
        AddLine(canvas, 10.5, 3.8, 14.5, 7.8, 1.4);
        AddLine(canvas, 10.5, 7.8, 14.5, 7.8, 1.4);
        var band = new XamlRectangle
        {
            Width = 12,
            Height = 3.5,
            RadiusX = 1.4,
            RadiusY = 1.4,
            Fill = new SolidColorBrush(Color.FromArgb(190, 255, 255, 255)),
            StrokeThickness = 0,
        };
        Canvas.SetLeft(band, 4);
        Canvas.SetTop(band, 9);
        canvas.Children.Add(band);
        AddLine(canvas, 6, 9.5, 12, 9.5, 1.3);
        AddLine(canvas, 6, 12, 11, 12, 1.3);
        return canvas;
    }

    private static FrameworkElement CreateFallbackIcon()
    {
        return new XamlEllipse
        {
            Width = 10,
            Height = 10,
            Fill = new SolidColorBrush(Colors.White),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
    }

    private void SetMenuTitle(MenuBarItem item, string title) => item.Title = title;

    private void SetMenuTooltip(FrameworkElement element, string tooltip)
    {
        ToolTipService.SetToolTip(element, tooltip);
        AutomationProperties.SetName(element, tooltip);
    }

    private void SetMenuText(MenuFlyoutItem item, string text) => item.Text = text;

    private void SetMenuText(MenuFlyoutSubItem item, string text) => item.Text = text;

    private void SetText(TextBlock textBlock, string content) => textBlock.Text = content;

    private void AppMenuButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement element)
        {
            FlyoutBase.ShowAttachedFlyout(element);
        }
    }

    private void DocumentTitleEditButton_Click(object sender, RoutedEventArgs e)
    {
        DocumentTitleTextBox.Focus(FocusState.Programmatic);
        DocumentTitleTextBox.SelectAll();
    }

    private void ToolPickerButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement element)
        {
            FlyoutBase.ShowAttachedFlyout(element);
        }
    }

    private void SetSection(ShellSection section)
    {
        _currentSection = section;
        CaptureSection.Visibility = section == ShellSection.Capture ? Visibility.Visible : Visibility.Collapsed;
        ProcedureSection.Visibility = section == ShellSection.Procedure ? Visibility.Visible : Visibility.Collapsed;
        SettingsSection.Visibility = section == ShellSection.Settings ? Visibility.Visible : Visibility.Collapsed;
    }

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentSection == ShellSection.Settings)
        {
            ReturnToProcedureSection(false);
            return;
        }

        CaptureSettingsSnapshot();
        SetSection(ShellSection.Settings);
        SelectSettingsNavSection("General");
    }

    private void CaptureSettingsSnapshot()
    {
        _settingsSnapshot = App.Preferences.Clone();
    }

    private void ReturnToProcedureSection(bool restoreSnapshot)
    {
        if (restoreSnapshot && _settingsSnapshot is not null)
        {
            App.Preferences.CopyFrom(_settingsSnapshot);
            App.Preferences.Save();
            ApplyPreferencesToUi();
            UpdateSettingsSummary();
        }

        SetSection(ShellSection.Procedure);
        _settingsSnapshot = null;
    }

    private void SelectSettingsNavSection(string sectionKey)
    {
        SetSettingsNavButtonState(SettingsGeneralNavButton, sectionKey == "General");
        SetSettingsNavButtonState(SettingsCaptureNavButton, sectionKey == "Capture");
        SetSettingsNavButtonState(SettingsExportNavButton, sectionKey == "Export");
        SetSettingsNavButtonState(SettingsPublicationNavButton, sectionKey == "Publication");
        SetSettingsNavButtonState(SettingsShortcutsNavButton, sectionKey == "Shortcuts");
        SetSettingsNavButtonState(SettingsAboutNavButton, sectionKey == "About");

        SetText(SettingsPageTitleText, sectionKey switch
        {
            "General" => T("Général", "General"),
            "Capture" => T("Capture", "Capture"),
            "Export" => T("Export", "Export"),
            "Publication" => T("Publication", "Publication"),
            "Shortcuts" => T("Raccourcis", "Shortcuts"),
            "About" => T("À propos", "About"),
            _ => T("Général", "General")
        });

        UIElement? target = sectionKey switch
        {
            "General" => SettingsGeneralCard,
            "Capture" => SettingsCaptureCard,
            "Export" => ExportSection,
            "Publication" => PublishSection,
            "Shortcuts" => HelpShortcutsTitleText,
            "About" => SettingsAboutCard,
            _ => SettingsGeneralCard
        };

        target?.StartBringIntoView();
    }

    private void SetSettingsNavButtonState(Button button, bool isSelected)
    {
        button.Background = isSelected ? (Brush?)TryGetBrush("ToolSelectedBrush") ?? new SolidColorBrush(Colors.Transparent) : new SolidColorBrush(Colors.Transparent);
    }

    private void SettingsNavButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not string sectionKey)
        {
            return;
        }

        if (sectionKey == "Shortcuts")
        {
            SetSection(ShellSection.Capture);
            HelpShortcutsTitleText.StartBringIntoView();
            StatusText.Text = T("Aide ouverte.", "Help opened.");
            return;
        }

        SelectSettingsNavSection(sectionKey);
    }

    private void SettingsSaveButton_Click(object sender, RoutedEventArgs e)
    {
        App.Preferences.Save();
        CaptureSettingsSnapshot();
        StatusText.Text = T("Réglages enregistrés.", "Settings saved.");
        SetSection(ShellSection.Procedure);
    }

    private void SettingsCancelButton_Click(object sender, RoutedEventArgs e)
    {
        ReturnToProcedureSection(restoreSnapshot: true);
        StatusText.Text = T("Réglages annulés.", "Settings cancelled.");
    }

    private void FloatingToolRailDragThumb_DragDelta(object sender, DragDeltaEventArgs e)
    {
        var currentLeft = App.Preferences.FloatingToolRailLeft;
        var currentTop = App.Preferences.FloatingToolRailTop;
        ApplyFloatingToolRailPosition(currentLeft + e.HorizontalChange, currentTop + e.VerticalChange);
    }

    private void FloatingToolRailDragThumb_DragCompleted(object sender, DragCompletedEventArgs e)
    {
        ApplyFloatingToolRailPosition(App.Preferences.FloatingToolRailLeft, App.Preferences.FloatingToolRailTop);
        App.Preferences.Save();
    }

    private void ShowHelpSectionMenuItem_Click(object sender, RoutedEventArgs e)
    {
        SetSection(ShellSection.Capture);
        HelpShortcutsTitleText.StartBringIntoView();
        StatusText.Text = T("Aide ouverte.", "Help opened.");
    }

    private void ReturnToProcedureMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (_currentSection == ShellSection.Settings)
        {
            ReturnToProcedureSection(restoreSnapshot: false);
        }
        else
        {
            SetSection(ShellSection.Procedure);
        }

        StatusText.Text = T("Retour à la procédure.", "Back to the procedure.");
    }

    private void CaptureShortcutMenuItem_Click(object sender, RoutedEventArgs e)
    {
        SetSection(ShellSection.Capture);
        HelpShortcutsTitleText.StartBringIntoView();

        var description = sender is MenuFlyoutItem { Tag: string tag }
            ? tag switch
            {
                "PrintScreen" => T("Impr écran : copie tout l’écran dans le presse-papiers.", "Print Screen: copies the whole screen to the clipboard."),
                "AltPrintScreen" => T("Alt + Impr écran : copie la fenêtre active dans le presse-papiers.", "Alt + Print Screen: copies the active window to the clipboard."),
                "WinPrintScreen" => T("Win + Impr écran : enregistre la capture dans Images > Captures d’écran.", "Win + Print Screen: saves the capture to Pictures > Screenshots."),
                "WinShiftS" => T("Win + Shift + S : ouvre l’outil Capture d’écran.", "Win + Shift + S: opens Snipping Tool."),
                "FnWinSpace" => T("Fn + Win + Espace : alternative sur certains claviers sans touche Impr écran.", "Fn + Win + Space: alternate on some keyboards without a Print Screen key."),
                _ => T("Aide capture ouverte.", "Screenshot help opened.")
            }
            : T("Aide capture ouverte.", "Screenshot help opened.");

        StatusText.Text = description;
    }

    private void CollapseStepsButton_Click(object sender, RoutedEventArgs e)
    {
        App.Preferences.CollapseStepsPane = !App.Preferences.CollapseStepsPane;
        App.Preferences.Save();
        ApplyPreferencesToUi();
    }

    private void InspectorToggleButton_Click(object sender, RoutedEventArgs e)
    {
        _isInspectorPaneCollapsed = !_isInspectorPaneCollapsed;
        ApplyProcedurePaneLayout();
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

    private void DocumentTabsListView_DragItemsCompleted(object sender, DragItemsCompletedEventArgs args)
    {
        if (_pendingDragSnapshot is not null)
        {
            var currentState = CaptureProjectState();
            if (!AreProjectStatesEquivalent(_pendingDragSnapshot, currentState))
            {
                PushUndoState(_pendingDragSnapshot);
                if (_currentDocument is not null)
                {
                    MarkDocumentDirty(_currentDocument);
                }
            }

            _pendingDragSnapshot = null;
        }

        UpdateStatus();
    }

    private void DocumentTabsListView_DragItemsStarting(object sender, DragItemsStartingEventArgs args)
    {
        _pendingDragSnapshot = CaptureProjectState();
    }

    private void DuplicateDocumentButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: ScreenshotDocument document })
        {
            PushUndoState(CaptureProjectState());
            var duplicate = CloneDocument(document);
            var index = _documents.IndexOf(document);
            _documents.Insert(Math.Max(0, index + 1), duplicate);
            RefreshDocumentOrderMetadata();
            duplicate.IsDirty = true;
            DocumentTabsListView.SelectedItem = duplicate;
        }
    }

    private async void UndoButton_Click(object sender, RoutedEventArgs e) => await UndoProjectAsync();

    private async void RedoButton_Click(object sender, RoutedEventArgs e) => await RedoProjectAsync();

    private void ToolbarDocumentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || sender is not ComboBox combo || combo.SelectedItem is not ScreenshotDocument document)
        {
            return;
        }

        if (ReferenceEquals(document, _currentDocument))
        {
            return;
        }

        DocumentTabsListView.SelectedItem = document;
    }

    private async void ToolbarPaletteComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _isSyncingPaletteUi || sender is not ComboBox combo || combo.SelectedItem is not GradientPaletteDefinition palette || _currentDocument is null)
        {
            return;
        }

        if (ReferenceEquals(combo.SelectedItem, palette) && string.Equals(_currentDocument.SelectedPaletteKey, palette.Key, StringComparison.Ordinal))
        {
            return;
        }

        await ApplyPaletteSelectionAsync(palette.Key, _currentDocument.SelectedPaletteShadeIndex);
    }

    private void ZoomInButton_Click(object sender, RoutedEventArgs e) => SetZoom(_currentZoom + 0.1);

    private void ZoomOutButton_Click(object sender, RoutedEventArgs e) => SetZoom(_currentZoom - 0.1);

    private void ZoomResetButton_Click(object sender, RoutedEventArgs e) => SetZoom(1.0);

    private void FitToViewportButton_Click(object sender, RoutedEventArgs e)
    {
        var contentWidth = PreviewViewport.Width;
        var contentHeight = PreviewViewport.Height;
        var availableWidth = CanvasScrollViewer.ActualWidth;
        var availableHeight = CanvasScrollViewer.ActualHeight;

        if (contentWidth <= 0 || contentHeight <= 0 || availableWidth <= 0 || availableHeight <= 0)
        {
            SetZoom(1.0);
            return;
        }

        var fitZoom = Math.Min(availableWidth / contentWidth, availableHeight / contentHeight);
        SetZoom(fitZoom);
    }

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
            PushUndoState(CaptureProjectState());
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
            PushUndoState(CaptureProjectState());
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
        MarkDocumentDirty(_currentDocument);
    }

    private void AuthorTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        _currentDocument.Author = textBox.Text;
        MarkDocumentDirty(_currentDocument);
    }

    private void DocumentVersionTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        _currentDocument.DocumentVersion = textBox.Text;
        MarkDocumentDirty(_currentDocument);
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

        UpdateSettingsSummary();
    }

    private async void BrowseExportFolderButton_Click(object sender, RoutedEventArgs e)
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
        if (ExportFolderTextBox is not null)
        {
            ExportFolderTextBox.Text = folder.Path;
        }

        UpdateSettingsSummary();
        StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("Dossier d’export : {0}", "Export folder: {0}"), folder.Path);
    }

    private void ExportFormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || ExportFormatComboBox.SelectedItem is not ComboBoxItem { Tag: ExportFormatChoice format })
        {
            return;
        }

        App.Preferences.DefaultExportFormat = format;
        App.Preferences.Save();
        UpdateSettingsSummary();
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
        UpdateSettingsSummary();
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
        UpdateSettingsSummary();
    }

    private void AutoImportClipboardToggleSwitch_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isInitializingUi)
        {
            return;
        }

        App.Preferences.AutoImportClipboard = AutoImportClipboardToggleSwitch.IsOn;
        App.Preferences.Save();
        UpdateSettingsSummary();
    }

    private void OpacitySlider_ValueChanged(object sender, RangeBaseValueChangedEventArgs e)
    {
        if (_currentDocument is null || _isInitializingUi || _isApplyingDocumentState || _isSyncingOpacityUi)
        {
            return;
        }

        var value = sender is RangeBase rangeBase ? rangeBase.Value : OpacitySlider.Value;
        _isSyncingOpacityUi = true;
        try
        {
            if (!ReferenceEquals(sender, OpacitySlider))
            {
                OpacitySlider.Value = value;
            }

            if (!ReferenceEquals(sender, ToolbarOpacitySlider))
            {
                ToolbarOpacitySlider.Value = value;
            }

            PushUndoState(CaptureProjectState());
            _currentDocument.DefaultOpacity = value;
            MarkDocumentDirty(_currentDocument);
        }
        finally
        {
            _isSyncingOpacityUi = false;
        }

        ToolbarOpacityValueText.Text = string.Format(CultureInfo.CurrentCulture, "{0:0} %", value * 100);
    }

    private void SelectedOpacitySlider_ValueChanged(object sender, RangeBaseValueChangedEventArgs e)
    {
        if (_currentDocument is null || _isInitializingUi || _isApplyingDocumentState || _isSyncingAnnotationSelectionUi)
        {
            return;
        }

        PushUndoState(CaptureProjectState());
        if (_selectedAnnotation is not null)
        {
            _selectedAnnotation.Opacity = SelectedOpacitySlider.Value;
            RenderAnnotations(_currentDocument);
            SelectAnnotation(_selectedAnnotation);
            MarkDocumentDirty(_currentDocument);
            return;
        }

        _currentDocument.DefaultOpacity = SelectedOpacitySlider.Value;
        MarkDocumentDirty(_currentDocument);
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

        PushUndoState(CaptureProjectState());
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

    private void BringSelectedAnnotationForwardButton_Click(object sender, RoutedEventArgs e)
        => MoveSelectedAnnotationLayer(toFront: true);

    private void SendSelectedAnnotationBackwardButton_Click(object sender, RoutedEventArgs e)
        => MoveSelectedAnnotationLayer(toFront: false);

    private void MoveSelectedAnnotationLayer(bool toFront)
    {
        if (_currentDocument is null || _selectedAnnotation is null)
        {
            return;
        }

        var index = _currentDocument.Annotations.IndexOf(_selectedAnnotation);
        if (index < 0)
        {
            return;
        }

        var targetIndex = toFront ? _currentDocument.Annotations.Count - 1 : 0;
        if (index == targetIndex)
        {
            return;
        }

        PushUndoState(CaptureProjectState());
        _currentDocument.Annotations.RemoveAt(index);
        _currentDocument.Annotations.Insert(targetIndex, _selectedAnnotation);
        MarkDocumentDirty(_currentDocument);
        RenderAnnotations(_currentDocument);
        SelectAnnotation(_selectedAnnotation);
        UpdateAnnotationCount(_currentDocument);
    }

    private void StepMoreButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: ScreenshotDocument document })
        {
            return;
        }

        var flyout = new MenuFlyout();

        var moveUp = new MenuFlyoutItem { Text = T("Monter", "Move up") };
        moveUp.Click += (_, _) => MoveDocument(document, -1);
        flyout.Items.Add(moveUp);

        var moveDown = new MenuFlyoutItem { Text = T("Descendre", "Move down") };
        moveDown.Click += (_, _) => MoveDocument(document, 1);
        flyout.Items.Add(moveDown);

        var duplicate = new MenuFlyoutItem { Text = T("Dupliquer", "Duplicate") };
        duplicate.Click += (_, _) =>
        {
            PushUndoState(CaptureProjectState());
            var clone = CloneDocument(document);
            var index = _documents.IndexOf(document);
            _documents.Insert(Math.Max(0, index + 1), clone);
            RefreshDocumentOrderMetadata();
            clone.IsDirty = true;
            DocumentTabsListView.SelectedItem = clone;
        };
        flyout.Items.Add(duplicate);

        flyout.Items.Add(new MenuFlyoutSeparator());

        var delete = new MenuFlyoutItem { Text = T("Supprimer", "Delete") };
        delete.Click += async (_, _) => await TryCloseDocumentAsync(document);
        flyout.Items.Add(delete);

        flyout.ShowAt((FrameworkElement)sender);
    }

    private void AddLegendButton_Click(object sender, RoutedEventArgs e)
    {
        SelectTool(EditorTool.Sticker);
        StatusText.Text = T("Cliquez dans le canvas pour ajouter une gommette.", "Click the canvas to add a sticker.");
    }

    private async void DeleteSelectedAnnotationButton_Click(object sender, RoutedEventArgs e)
        => await DeleteSelectedAnnotationAsync();

    private async Task DeleteSelectedAnnotationAsync()
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

        PushUndoState(CaptureProjectState());
        _currentDocument.Annotations.Remove(selected);
        _selectedAnnotation = null;
        StopInlineTextEdit(rerender: false, refreshSelection: false);
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

    private void RootGrid_PreviewKeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == Windows.System.VirtualKey.Escape && _currentSection != ShellSection.Procedure)
        {
            if (e.OriginalSource is not TextBox)
            {
                SetSection(ShellSection.Procedure);
                StatusText.Text = T("Retour à la procédure.", "Back to the procedure.");
                e.Handled = true;
            }

            return;
        }

        if (e.Key != Windows.System.VirtualKey.Delete || _currentDocument is null || _selectedAnnotation is null)
        {
            return;
        }

        if (e.OriginalSource is TextBox)
        {
            return;
        }

        e.Handled = true;
        _ = DeleteSelectedAnnotationAsync();
    }

    private void BeginInlineTextEdit(AnnotationModel annotation)
    {
        if (_currentDocument is null || annotation.Kind != AnnotationKind.Text)
        {
            return;
        }

        _editingTextAnnotation = annotation;
        RenderAnnotations(_currentDocument);
        SelectAnnotation(annotation);
        StatusText.Text = T("Texte modifiable dans l’aperçu.", "Text editable in the preview.");
    }

    private void StopInlineTextEdit(bool rerender = true, bool refreshSelection = true)
    {
        if (_editingTextAnnotation is null)
        {
            return;
        }

        _editingTextAnnotation = null;
        if (!rerender || _currentDocument is null)
        {
            UpdateStatus();
            return;
        }

        RenderAnnotations(_currentDocument);
        if (!refreshSelection)
        {
            UpdateStatus();
            return;
        }

        if (_selectedAnnotation is not null)
        {
            SelectAnnotation(_selectedAnnotation);
        }
        else
        {
            SelectAnnotation(null);
        }

        UpdateStatus();
    }

    private void InlineTextEditor_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_currentDocument is null || sender is not TextBox editor || editor.Tag is not AnnotationModel annotation)
        {
            return;
        }

        if (_isApplyingDocumentState)
        {
            return;
        }

        annotation.Text = editor.Text;
        var size = MeasureTextAnnotationSize(annotation.Text, annotation.FontSize);
        annotation.Bounds = new Rect(annotation.Bounds.X, annotation.Bounds.Y, size.Width, size.Height);
        if (editor.Parent is Border border)
        {
            border.Width = size.Width;
            border.Height = size.Height;
        }

        UpdateSelectionOutline(annotation);
        if (ReferenceEquals(_selectedAnnotation, annotation))
        {
            _isSyncingAnnotationSelectionUi = true;
            try
            {
                AnnotationTextBox.Text = editor.Text;
            }
            finally
            {
                _isSyncingAnnotationSelectionUi = false;
            }

            SelectedAnnotationSummaryText.Text = annotation.Kind switch
            {
                AnnotationKind.Sticker or AnnotationKind.Text => $"{annotation.Text} · {annotation.Opacity:P0}",
                AnnotationKind.Rectangle or AnnotationKind.Ellipse => $"{annotation.Bounds.Width:0} x {annotation.Bounds.Height:0} · {annotation.Opacity:P0}",
                _ => annotation.Opacity.ToString("P0", CultureInfo.CurrentCulture)
            };
        }

        MarkDocumentDirty(_currentDocument);
        UpdateAnnotationCount(_currentDocument);
        UpdateLegendList();
    }

    private void InlineTextEditor_LostFocus(object sender, RoutedEventArgs e)
    {
        if (_editingTextAnnotation is null)
        {
            return;
        }

        StopInlineTextEdit();
    }

    private void AnnotationElement_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
    {
        if (sender is FrameworkElement { Tag: AnnotationModel annotation } && annotation.Kind == AnnotationKind.Text)
        {
            BeginInlineTextEdit(annotation);
            e.Handled = true;
        }
    }

    private async void ExportMarkdownButton_Click(object sender, RoutedEventArgs e) => await ExportGuideAsync(ExportFormatChoice.Markdown);

    private async void ExportHtmlButton_Click(object sender, RoutedEventArgs e) => await ExportGuideAsync(ExportFormatChoice.Html);

    private async void ExportPdfButton_Click(object sender, RoutedEventArgs e) => await ExportGuideAsync(ExportFormatChoice.Pdf);

    private async void ExportDocxButton_Click(object sender, RoutedEventArgs e) => await ExportGuideAsync(ExportFormatChoice.Docx);

    private async void ExportAllFormatsButton_Click(object sender, RoutedEventArgs e) =>
        await ExportGuideAsync(ExportFormatChoice.Markdown, ExportFormatChoice.Html, ExportFormatChoice.Pdf, ExportFormatChoice.Docx);

    private async void ExportCurrentMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null)
        {
            return;
        }

        await SaveDocumentAsync(_currentDocument);
    }

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

        PushUndoState(CaptureProjectState());
        _documents.Move(index, targetIndex);
        RefreshDocumentOrderMetadata();
        document.IsDirty = true;
        DocumentTabsListView.SelectedItem = document;
        UpdateStatus();
    }

    private void SyncDocumentEditors(ScreenshotDocument document)
    {
        DocumentTitleTextBox.Text = document.BaseTitle;
        ToolbarDocumentComboBox.SelectedItem = document;
        _isSyncingPaletteUi = true;
        try
        {
            ToolbarPaletteComboBox.SelectedItem = _paletteOrder.FirstOrDefault(entry => string.Equals(entry.Key, document.SelectedPaletteKey, StringComparison.Ordinal));
        }
        finally
        {
            _isSyncingPaletteUi = false;
        }
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
        ToolbarOpacitySlider.Value = document.DefaultOpacity;
        ToolbarOpacityValueText.Text = string.Format(CultureInfo.CurrentCulture, "{0:0} %", document.DefaultOpacity * 100);
        SelectedOpacitySlider.Value = _selectedAnnotation?.Opacity ?? document.DefaultOpacity;
        ResetStickerNumberOnColorChangeCheckBox.IsChecked = document.ResetStickerNumberOnColorChange;
        ExportTemplateComboBox.SelectedIndex = (int)document.TemplateType;
        ExportFolderTextBox.Text = App.Preferences.ExportFolder;
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
        LegendPanel.Visibility = _selectedAnnotation is null || _selectedAnnotation.Kind == AnnotationKind.Sticker
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    private void SelectAnnotation(AnnotationModel? annotation)
    {
        if (_editingTextAnnotation is not null && !ReferenceEquals(_editingTextAnnotation, annotation))
        {
            StopInlineTextEdit(refreshSelection: false);
        }

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
                StepDetailsSubtitleText.Text = T("Titre et note de l’étape", "Step title and note");
                SelectedAnnotationSummaryText.Text = T("Aucun objet sélectionné.", "No object selected.");
                SelectedColorLabelText.Visibility = Visibility.Collapsed;
                SelectedColorText.Visibility = Visibility.Collapsed;
                SelectedOpacitySlider.Value = _currentDocument.DefaultOpacity;
                return;
            }

            StepDetailsPanel.Visibility = Visibility.Collapsed;
            AnnotationContextPanel.Visibility = Visibility.Visible;
            StepDetailsSubtitleText.Text = T("Objet sélectionné", "Selected object");
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
            var paletteName = _palettes.TryGetValue(annotation.PaletteKey, out var selectedPalette)
                ? selectedPalette.DisplayName
                : annotation.PaletteKey;

            SelectedColorLabelText.Visibility = Visibility.Visible;
            SelectedColorText.Visibility = Visibility.Visible;
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
            SelectedColorLabelText.Text = T("Couleur", "Color");
            SelectedColorText.Text = string.Format(CultureInfo.CurrentCulture, "{0} · {1}", paletteName, annotation.PaletteShadeIndex + 1);
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

        UpdateDraggedAnnotation(e.GetCurrentPoint(SceneHost).Position);
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
        _annotationDragStart = e.GetCurrentPoint(SceneHost).Position;
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

    private static string NormalizeLineBreaks(string value)
    {
        return (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
    }

    private static string FormatMarkdownMultiline(string value)
    {
        return string.Join("  \n", NormalizeLineBreaks(value).Split('\n'));
    }

    private static string FormatHtmlMultiline(string value)
    {
        return EscapeHtml(NormalizeLineBreaks(value)).Replace("\n", "<br />");
    }

    private string BuildMarkdown(GuideExportManifest manifest, bool includeImageReferences = true)
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
                markdown.AppendLine(FormatMarkdownMultiline(step.Note));
                markdown.AppendLine();
            }

            if (includeImageReferences)
            {
                markdown.AppendLine($"![](images/{Path.GetFileName(step.ImagePath)})");
                markdown.AppendLine();
            }

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
                html.AppendLine($"<p>{FormatHtmlMultiline(step.Note)}</p>");
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
