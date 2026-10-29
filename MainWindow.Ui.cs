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
        EditingPreferences.Language == LanguageChoice.English ||
        (EditingPreferences.Language == LanguageChoice.System &&
         CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.Equals("en", StringComparison.OrdinalIgnoreCase));


    private sealed record SettingsShortcutItem(string Key, string Description);

    private void PopulatePaletteOrder()
    {
        _paletteOrder.Clear();
        foreach (var key in new[] { "Sunset", "Ember", "Citrus", "Mint", "Lagoon", "Sky", "Ocean", "Rose", "Berry", "Grape" })
        {
            if (_palettes.TryGetValue(key, out var palette))
            {
                _paletteOrder.Add(palette with { DisplayName = SafetyText("Palette" + palette.Key) });
            }
        }
    }

    private void ApplyChromeLabels()
    {
        LocalizedText.CurrentLanguage = UseEnglish ? "en-US" : "fr-FR";
        LocalizePalettes();
        RefreshAnnotationObjects();
        ImportClipboardMenuItem.Text = SafetyText("ImportClipboard.Text");
        ExportCancelButton.Content = SafetyText("ExportCancel.Content");
        if (_currentDocument is not null) { CurrentToolText.Text = ToolbarCurrentToolText.Text = SafetyText(_toolLabels[_currentTool]); }
        SetMenuTitle(FileMenuBarItem, SafetyText("UiFile"));
        SetMenuTitle(EditMenuBarItem, SafetyText("UiEdit"));
        SetMenuTitle(ImageMenuBarItem, SafetyText("UiImage"));
        SetMenuTitle(ExportMenuBarItem, SafetyText("UiExport"));
        SetMenuTitle(HelpMenuBarItem, SafetyText("UiHelp"));
        SetMenuTooltip(FileMenuBarItem, SafetyText("UiOpenFileActions"));
        SetMenuTooltip(EditMenuBarItem, SafetyText("UiOpenEditActions"));
        SetMenuTooltip(ImageMenuBarItem, SafetyText("UiOpenImageActions"));
        SetMenuTooltip(ExportMenuBarItem, SafetyText("UiOpenExportOptions"));
        SetMenuTooltip(HelpMenuBarItem, SafetyText("UiOpenHelpAndShortcuts"));

        SetMenuText(NewProjectMenuItem, SafetyText("UiNew"));
        SetMenuText(OpenProjectMenuItem, SafetyText("UiOpen"));
        SetMenuText(SaveProjectMenuItem, SafetyText("UiSave"));
        SetMenuText(UndoMenuItem, SafetyText("UiUndo"));
        SetMenuText(RedoMenuItem, SafetyText("UiRedo"));
        SetMenuText(CopyImageMenuItem, SafetyText("ClipboardCopyHelp"));
        SetButtonContent(CopyImageButton, SafetyText("ClipboardCopy"));
        SetButtonTooltip(CopyImageButton, SafetyText("ClipboardCopyHelp"));
        SetMenuText(ImportScreenshotMenuItem, SafetyText("UiImport"));
        SetMenuText(OcrMenuItem, "OCR");
        SetMenuText(ExportCurrentMenuItem, SafetyText("UiCurrentStepPNG"));
        SetMenuText(ExportAllFormatsMenuItem, SafetyText("ExportGuideFormats"));
        SetMenuText(ExportMarkdownMenuItem, "Markdown");
        SetMenuText(ExportHtmlMenuItem, "HTML");
        SetMenuText(ExportPdfMenuItem, "PDF");
        SetMenuText(ExportDocxMenuItem, SafetyText("UiWord"));
        SetMenuText(PublishNotionMenuItem, SafetyText("UiToNotionTextOnly"));
        SetMenuText(PublishConfluenceMenuItem, SafetyText("UiToConfluenceTextOnly"));
        SetMenuText(PublishSharePointMenuItem, SafetyText("UiToSharePoint"));
        SetMenuText(OpenHelpMenuItem, SafetyText("UiQuickHelp"));
        SetMenuText(PrintScreenMenuItem, SafetyText("UiPrintScreen"));
        SetMenuText(AltPrintScreenMenuItem, SafetyText("UiAltPrintScreen"));
        SetMenuText(WinPrintScreenMenuItem, SafetyText("UiWinPrintScreen"));
        SetMenuText(WinShiftSMenuItem, SafetyText("UiWinShiftS"));
        SetMenuText(FnWinSpaceMenuItem, SafetyText("UiFnWinSpace"));
        SetButtonContent(SettingsMenuButton, SafetyText("UiSettings"));
        SetButtonTooltip(SettingsMenuButton, SafetyText("UiOpenSettings"));
        SetButtonTooltip(AppMenuButton, SafetyText("UiApplicationMenu"));
        SetButtonTooltip(DocumentTitleEditButton, SafetyText("UiEditTheDocumentTitle"));
        SetButtonContent(InspectorToggleButton, SafetyText("UiText"));
        SetButtonTooltip(InspectorToggleButton, SafetyText("UiShowOrHideTheInspector"));

        SetText(SettingsGeneralNavText, SafetyText("UiGeneral"));
        SetText(SettingsCaptureNavText, SafetyText("UiCapture"));
        SetText(SettingsExportNavText, SafetyText("UiExport"));
        SetText(SettingsPublicationNavText, SafetyText("UiPublication"));
        SetText(SettingsShortcutsNavText, SafetyText("UiShortcuts"));
        SetText(SettingsAboutNavText, SafetyText("UiAbout"));

        SetCommandButton(NewProjectButton, SafetyText("UiNew"), SafetyText("UiNewProject"));
        SetCommandButton(OpenProjectButton, SafetyText("UiOpen"), SafetyText("UiOpenAProject_207d0454"));
        SetCommandButton(SaveProjectButton, SafetyText("UiSave"), SafetyText("UiSaveProject"));
        SetCommandButton(DemoProjectButton, SafetyText("UiDemo"), SafetyText("UiLoadTheDemoTemplate_d853f526"));
        SetButtonContent(CaptureQuickImportButton, SafetyText("UiImport"));
        SetButtonContent(ImportScreenshotButton, SafetyText("UiImport"));
        SetCommandButton(OcrButton, SafetyText("UiOCR"), SafetyText("UiRecognizeTextFromTheCurrentImage"));
        SetButtonContent(ExportCurrentButton, SafetyText("UiExport_014c6f83"));
        SetCommandButton(UndoButton, SafetyText("UiUndo"), SafetyText("UiUndoTheLastAction"));
        SetCommandButton(RedoButton, SafetyText("UiRedo_2758230a"), SafetyText("UiRedoTheLastAction"));
        SetButtonContent(CollapseStepsButton, SafetyText("UiHideSteps"));
        SetButtonContent(BrowseExportFolderButton, SafetyText("UiBrowse"));
        SetButtonContent(BrowseSharePointFolderButton, SafetyText("UiBrowse"));
        SetButtonContent(ZoomOutButton, SafetyText("UiText_ffb392d1"));
        SetButtonContent(ZoomResetButton, SafetyText("Ui100"));
        SetButtonContent(ZoomInButton, SafetyText("UiText_e63541c0"));
        ZoomLevelText.Text = string.Format(CultureInfo.CurrentCulture, "{0:0} %", _currentZoom * 100);
        SetButtonContent(ExportPngButton, "PNG");
        SetButtonContent(ExportAllFormatsButton, SafetyText("UiAll"));
        SetButtonContent(ExportMarkdownButton, "Markdown");
        SetButtonContent(ExportHtmlButton, "HTML");
        SetButtonContent(ExportPdfButton, "PDF");
        SetButtonContent(ExportDocxButton, SafetyText("UiWord"));
        SetCommandButton(ApplyCropButton, SafetyText("UiApply"), SafetyText("UiApplyCrop"));
        SetCommandButton(ResetCropButton, SafetyText("UiReset"), SafetyText("UiResetCrop"));
        SetCommandButton(ClearAnnotationsButton, SafetyText("UiClearAll"), SafetyText("UiClearAllAnnotations"));
        SetCommandButton(DuplicateSelectedButton, SafetyText("UiDuplicate"), SafetyText("UiDuplicateSelectedObject"));
        SetCommandButton(DeleteSelectedButton, SafetyText("UiDelete"), SafetyText("UiDeleteSelectedObject"));
        SetCommandButton(BringForwardButton, SafetyText("UiFront"), SafetyText("UiBringTheObjectToFront"));
        SetCommandButton(SendBackwardButton, SafetyText("UiBack"), SafetyText("UiSendTheObjectToBack"));
        SetCommandButton(AddLegendButton, SafetyText("UiLegend"), SafetyText("UiAddALegend"));
        SetCommandButton(PublishNotionButton, SafetyText("UiPublishTextToNotion"), SafetyText("PublicationTextOnlyNotice"));
        SetCommandButton(PublishConfluenceButton, SafetyText("UiPublishTextToConfluenceCloud"), SafetyText("PublicationTextOnlyNotice"));
        SetCommandButton(PublishSharePointButton, SafetyText("UiPublishToSharePoint"), SafetyText("UiCopyThePackageToTheSynchronizedFolder"));
        SetButtonContent(AddStepButton, SafetyText("UiAdd"));
        SetButtonContent(CollapseStepsButton, SafetyText("UiText_7c12212b"));
        SetButtonTooltip(CollapseStepsButton, SafetyText("UiHideTheSteps"));

        SetCommandButton(ToolbarZoomOutButton, SafetyText("UiText_ffb392d1"), SafetyText("UiZoomOut"));
        SetCommandButton(ToolbarZoomResetButton, SafetyText("Ui100"), SafetyText("UiResetTo100"));
        SetCommandButton(ToolbarZoomInButton, SafetyText("UiText_e63541c0"), SafetyText("UiZoomIn"));
        SetCommandButton(ToolbarFitButton, SafetyText("UiFit"), SafetyText("UiFitToView"));
        SetCommandButton(ToolbarStepsButton, SafetyText("UiSteps"), SafetyText("UiShowOrHideTheStepsPane"));

        SetToolButton(ToolbarSelectButton, EditorTool.Select, SafetyText("UiSelect"));
        SetToolButton(ToolbarCropButton, EditorTool.Crop, SafetyText("UiCropTheCapture"));
        SetOcrButton(ToolbarOcrButton, SafetyText("UiOCR"));
        SetToolButton(ToolbarTextButton, EditorTool.Text, SafetyText("UiText_0420e350"));
        SetToolButton(ToolbarStickerButton, EditorTool.Sticker, SafetyText("UiSticker"));
        SetToolButton(ToolbarArrowStraightButton, EditorTool.ArrowStraight, SafetyText("UiStraightArrow"));
        SetToolButton(ToolbarArrowCurvedButton, EditorTool.ArrowCurved, SafetyText("UiCurvedArrow"));
        SetToolButton(ToolbarRectangleButton, EditorTool.Rectangle, SafetyText("UiRectangle"));
        SetToolButton(ToolbarEllipseButton, EditorTool.Ellipse, SafetyText("UiEllipse"));
        SetToolButton(ToolbarFocusButton, EditorTool.Focus, SafetyText("UiFocus"));
        SetToolButton(ToolbarMaskButton, EditorTool.Mask, SafetyText("UiMask"));
        SetButtonTooltip(ToolbarMaskButton, SafetyText("MaskHelp"));
        SetToolButton(ToolbarToolPickerButton, _currentTool, SafetyText("UiChooseATool"));

        SetText(CapturePageTitleText, SafetyText("UiHelp"));
        SetText(CapturePageSummaryText, SafetyText("UiWindowsCaptureShortcutsAndUsefulSnapSlateShortcuts"));
        SetText(CaptureHowTitleText, SafetyText("UiHowItWorks"));
        SetText(CaptureHowBodyText, SafetyText("UiPrintScreenAltPrintScreenWinPrint"));
        SetText(CaptureStepOneTitleText, SafetyText("UiPrintScreen"));
        SetText(CaptureStepOneBodyText, SafetyText("UiCopiesTheWholeScreenToTheClipboard"));
        SetText(CaptureStepTwoTitleText, SafetyText("UiAltPrintScreen"));
        SetText(CaptureStepTwoBodyText, SafetyText("UiCopiesOnlyTheActiveWindow"));
        SetText(CaptureStepThreeTitleText, SafetyText("UiWinPrintScreen"));
        SetText(CaptureStepThreeBodyText, SafetyText("UiSavesTheCaptureToPicturesScreenshots"));
        SetText(CaptureShortcutsTitleText, SafetyText("UiCaptureShortcuts"));
        SetText(CaptureShortcutOneText, SafetyText("UiWinShiftSSnippingToolWithFreeform"));
        SetText(CaptureShortcutTwoText, SafetyText("UiFnWinSpaceUsefulOnKeyboardsWithout"));
        SetText(CaptureShortcutThreeText, SafetyText("UiPrtScnAltPrtScnAndWinPrtScnWork"));
        SetText(CaptureShortcutFourText, SafetyText("UiImportedScreenshotsCanAlsoArriveAutomaticallyFrom"));
        SetText(CaptureShortcutFiveText, SafetyText("UiCtrlOImportAnImageFile"));
        SetText(HelpShortcutsTitleText, SafetyText("UiUsefulShortcuts"));
        SetText(CaptureStatusTitleText, SafetyText("UiStatus"));
        SetText(CaptureStatusBodyText, SafetyText("UiAutomaticImportWorksWhileSnapSlateIsOpen"));
        SetText(ProcedureTitleLabelText, SafetyText("UiDocument"));
        SetText(ToolbarDocumentLabelText, SafetyText("UiDocument"));
        SetText(CurrentToolLabelText, SafetyText("UiTool"));
        SetText(ToolbarCurrentToolLabelText, SafetyText("UiTool_9ad395ff"));
        SetText(StepsPaneHintText, SafetyText("UiDragToReorderTheSteps"));
        SetText(PaletteSectionTitleText, SafetyText("UiPalette"));
        SetText(ToolbarPaletteLabelText, SafetyText("UiPalette"));
        SetText(ShadeSectionTitleText, SafetyText("UiShades"));
        SetText(ToolbarOpacityLabelText, SafetyText("UiOpacity"));
        SetText(ToolbarOpacityValueText, "100 %");
        SetText(StepsPaneTitleText, SafetyText("UiSteps"));
        SetText(StepDetailsSubtitleText, SafetyText("UiStepTitleAndNote"));
        SetText(SelectedColorLabelText, SafetyText("UiColor"));
        CanvasSummaryTitleText.Visibility = Visibility.Collapsed;
        CanvasSummaryText.Visibility = Visibility.Collapsed;
        CanvasFileLabelText.Visibility = Visibility.Collapsed;
        CanvasFileNameText.Visibility = Visibility.Collapsed;
        SetText(CanvasPreviewLabelText, SafetyText("UiPreview"));
        SetText(StepDetailsTitleText, SafetyText("UiStepDetails"));
        SetText(StepTitleLabelText, SafetyText("UiStepTitle"));
        SetText(StepNoteLabelText, SafetyText("UiStepNote"));
        SetText(TemplateTypeLabelText, SafetyText("UiDocumentType"));
        SetText(AudienceLabelText, SafetyText("UiAudience"));
        SetText(AuthorLabelText, SafetyText("UiAuthor"));
        SetText(DocumentVersionLabelText, SafetyText("UiVersion"));
        SetText(AnnotationPanelTitleText, SafetyText("UiSelectedObject"));
        SetText(StickerValueLabelText, SafetyText("UiStickerValue"));
        SetText(StickerLegendLabelText, SafetyText("UiStickerLegend"));
        SetText(AnnotationTextLabelText, SafetyText("UiText_0420e350"));
        SetText(StickerModeLabelText, SafetyText("UiStickerFormat"));
        SetText(StrokeThicknessLabelText, SafetyText("UiThickness"));
        SetText(SelectedOpacityLabelText, SafetyText("UiOpacity"));
        SetText(LegendTitleText, SafetyText("UiStickerLegends"));
        SetText(ExportPageTitleText, SafetyText("UiExport"));
        SetText(ExportPageSummaryText, SafetyText("ExportPreferencesHelp"));
        SetText(ExportTemplateLabelText, SafetyText("UiDocumentType"));
        SetText(ExportFormatLabelText, SafetyText("UiMainFormat"));
        SetText(ExportAudienceLabelText, SafetyText("UiAudience_4bfa6ed2"));
        SetText(ExportAuthorLabelText, SafetyText("UiAuthor"));
        SetText(ExportVersionLabelText, SafetyText("UiVersion"));
        SetText(ExportFolderLabelText, SafetyText("UiExportFolder"));
        SetText(ExportTemplateValueText, SafetyText("UiCurrentTypeUserManual"));
        SetText(ExportPreviewTitleText, SafetyText("UiSummary"));
        SetText(ExportPreviewSummaryText, SafetyText("UiAllGeneratesMarkdownHTMLPDFAndDOCX"));
        SetText(PublishPageTitleText, SafetyText("UiPublish"));
        SetText(PublishPageSummaryText, SafetyText("PublicationModes"));
        SetText(OcrSectionTitleText, SafetyText("UiOCR"));
        SetText(OcrSectionHelpText, SafetyText("UiRecognizedTextIsInsertedAsAText"));
        SetText(NotionSectionTitleText, SafetyText("UiNotion"));
        SetText(NotionSectionHelpText, SafetyText("UiShareTheParentPageWithYourNotion"));
        SetText(NotionParentPageLabelText, SafetyText("UiParentPageID"));
        SetText(NotionTokenLabelText, SafetyText("UiNotionAPIToken"));
        SetText(ConfluenceSectionTitleText, SafetyText("UiConfluence"));
        SetText(ConfluenceSectionHelpText, SafetyText("UiCreateAPageInAConfluenceCloud"));
        SetText(ConfluenceBaseUrlLabelText, SafetyText("UiSiteURL"));
        SetText(ConfluenceSpaceKeyLabelText, SafetyText("UiSpaceKey"));
        SetText(ConfluenceParentPageLabelText, SafetyText("UiParentPageIDOptional"));
        SetText(ConfluenceEmailLabelText, SafetyText("UiAccountEmail"));
        SetText(ConfluenceTokenLabelText, SafetyText("UiConfluenceAPIToken"));
        SetText(SharePointSectionTitleText, SafetyText("UiSharePoint"));
        SetText(SharePointSectionHelpText, SafetyText("UiCopyTheExportPackageIntoAFolder"));
        SetText(SharePointFolderLabelText, SafetyText("UiSynchronizedSharePointFolder"));
        SetText(SettingsPageTitleText, SafetyText("UiGeneral"));
        SetText(AppearanceSettingsTitleText, SafetyText("UiAppearance"));
        SetText(LanguageSettingsTitleText, SafetyText("UiLanguage"));
        SetText(CaptureSettingsTitleText, SafetyText("UiCapture"));
        SetText(ThemeSettingLabelText, SafetyText("UiTheme"));
        SetText(LanguageSettingLabelText, SafetyText("UiAppLanguage"));
        SetText(CaptureSettingLabelText, SafetyText("UiAutomaticImport"));
        SetText(ThemeSettingHelpText, SafetyText("UiSystemFollowsWindowsAutomatically"));
        SetText(LanguageSettingHelpText, SafetyText("UiByDefaultSnapSlateFollowsTheWindowsLanguage"));
        SetText(CaptureSettingHelpText, SafetyText("UiEnablesAutomaticImportOfCapturesCopiedTo"));
        SetText(SettingsRestartHintText, SafetyText("SettingsPreviewHelp"));
        SetText(SettingsAboutTitleText, SafetyText("UiAbout"));
        SetText(SettingsAboutBodyText, SafetyText("UiSnapSlateAcceleratesVisualProcedureCreationWithCaptures"));
        SetButtonContent(SettingsCancelButton, SafetyText("UiCancel"));
        SetButtonContent(SettingsSaveButton, SafetyText("UiSave_63e08f48"));
        SetText(SettingsTipTitleText, SafetyText("UiSnapSlateTip"));
        SetText(SettingsTipBodyText, SafetyText("UiUsefulShortcutsAreGroupedInHelpSo"));
        AutoImportClipboardToggleSwitch.OnContent = string.Empty;
        AutoImportClipboardToggleSwitch.OffContent = string.Empty;
        ApplyWorkspaceChrome();
    }

    private void UpdateExportSummaryList()
    {
        ExportStepsListView.ItemsSource = new[]
        {
            SafetyText("UiCurrentStepPNG"),
            SafetyText("UiMarkdownImages"),
            SafetyText("UiHTMLImages"),
            SafetyText("UiDocumentPDF"),
            SafetyText("UiWordDOCX")
        };
        HelpShortcutsListView.ItemsSource = new[]
        {
            new SettingsShortcutItem("Ctrl + Z", SafetyText("UiUndoTheLastAction")),
            new SettingsShortcutItem("Ctrl + Y", SafetyText("UiRedoTheLastAction")),
            new SettingsShortcutItem("Ctrl + S", SafetyText("UiSaveTheProject")),
            new SettingsShortcutItem("Ctrl + O", SafetyText("UiOpenAProject_207d0454")),
            new SettingsShortcutItem("Ctrl + N", SafetyText("UiNewProject")),
            new SettingsShortcutItem("Ctrl + Shift + N", SafetyText("UiDemoTemplate")),
            new SettingsShortcutItem("Win + Shift + S", SafetyText("CaptureShortcutHelp")),
            new SettingsShortcutItem("Ctrl + Shift + V", SafetyText("ImportClipboard.Text")),
            new SettingsShortcutItem("Ctrl + Shift + C", SafetyText("ClipboardCopy")),
            new SettingsShortcutItem("V / T / A / R / O / M", SafetyText("ToolShortcuts")),
            new SettingsShortcutItem("Suppr", SafetyText("UiDeleteTheSelectedObject"))
        };
    }

    private string GetThemeDisplayName(ThemeChoice theme)
    {
        return theme switch
        {
            ThemeChoice.Light => SafetyText("UiLight"),
            ThemeChoice.Dark => SafetyText("UiDark"),
            _ => SafetyText("UiSystem")
        };
    }

    private string GetLanguageDisplayName(LanguageChoice language)
    {
        return language switch
        {
            LanguageChoice.French => SafetyText("UiFrench"),
            LanguageChoice.English => SafetyText("UiEnglish"),
            _ => SafetyText("UiFollowWindows")
        };
    }

    private string GetExportFormatDisplayName(ExportFormatChoice format)
    {
        return format switch
        {
            ExportFormatChoice.Html => "HTML",
            ExportFormatChoice.Markdown => "Markdown",
            ExportFormatChoice.Pdf => "PDF",
            ExportFormatChoice.Docx => SafetyText("UiWord"),
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
            GuideTemplateType.UserManual => SafetyText("UiUserManual"),
            GuideTemplateType.InternalProcedure => SafetyText("UiInternalProcedure"),
            GuideTemplateType.SupportPlaybook => SafetyText("UiSupportPlaybook"),
            GuideTemplateType.ReleaseNotes => SafetyText("UiReleaseNotes"),
            _ => SafetyText("UiUserManual")
        };
    }

    private void UpdateExportTemplateSummary()
    {
        if (ExportTemplateValueText is null)
        {
            return;
        }

        var templateType = EditingMetadata.TemplateType;
        ExportTemplateValueText.Text = string.Format(
            CultureInfo.CurrentCulture,
            SafetyText("UiCurrentType0"),
            GetTemplateDisplayName(templateType));
    }

    private void ConfigureComboBoxes()
    {
        PopulateComboBox(ThemeComboBox,
            (SafetyText("UiSystem"), ThemeChoice.System),
            (SafetyText("UiLight"), ThemeChoice.Light),
            (SafetyText("UiDark"), ThemeChoice.Dark));

        PopulateComboBox(LanguageComboBox,
            (SafetyText("UiFollowWindows"), LanguageChoice.System),
            (SafetyText("UiFrench"), LanguageChoice.French),
            (SafetyText("UiEnglish"), LanguageChoice.English));

        PopulateComboBox(ExportFormatComboBox,
            ("PNG", ExportFormatChoice.Png),
            ("Markdown", ExportFormatChoice.Markdown),
            ("HTML", ExportFormatChoice.Html),
            ("PDF", ExportFormatChoice.Pdf),
            (SafetyText("UiWord"), ExportFormatChoice.Docx));

        PopulateComboBox(TemplateTypeComboBox,
            (SafetyText("UiUserManual"), GuideTemplateType.UserManual),
            (SafetyText("UiInternalProcedure"), GuideTemplateType.InternalProcedure),
            (SafetyText("UiSupportPlaybook"), GuideTemplateType.SupportPlaybook),
            (SafetyText("UiReleaseNotes"), GuideTemplateType.ReleaseNotes));

        PopulateComboBox(ExportTemplateComboBox,
            (SafetyText("UiUserManual"), GuideTemplateType.UserManual),
            (SafetyText("UiInternalProcedure"), GuideTemplateType.InternalProcedure),
            (SafetyText("UiSupportPlaybook"), GuideTemplateType.SupportPlaybook),
            (SafetyText("UiReleaseNotes"), GuideTemplateType.ReleaseNotes));

        TemplateTypeComboBox.SelectedIndex = 0;
        ExportTemplateComboBox.SelectedIndex = 0;

        PopulateComboBox(StickerModeComboBox,
            (SafetyText("UiNumeric"), 0),
            (SafetyText("UiAlphabetic"), 1));
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
        RootGrid.RequestedTheme = EditingPreferences.GetRequestedTheme();
        ThemeComboBox.SelectedIndex = (int)EditingPreferences.Theme;
        LanguageComboBox.SelectedIndex = (int)EditingPreferences.Language;
        ExportFormatComboBox.SelectedIndex = (int)EditingPreferences.DefaultExportFormat;
        AutoImportClipboardToggleSwitch.IsOn = EditingPreferences.AutoImportClipboard;
        ExportFolderTextBox.Text = EditingPreferences.ExportFolder;
        NotionParentPageIdTextBox.Text = EditingPreferences.NotionParentPageId;
        ConfluenceBaseUrlTextBox.Text = EditingPreferences.ConfluenceBaseUrl;
        ConfluenceSpaceKeyTextBox.Text = EditingPreferences.ConfluenceSpaceKey;
        ConfluenceParentPageIdTextBox.Text = EditingPreferences.ConfluenceParentPageId;
        ConfluenceEmailTextBox.Text = EditingPreferences.ConfluenceEmail;
        SharePointSyncFolderTextBox.Text = EditingPreferences.SharePointSyncFolder;
        ApplyProcedurePaneLayout();
        ApplyFloatingToolRailPosition(EditingPreferences.FloatingToolRailLeft, EditingPreferences.FloatingToolRailTop);
        UpdateSettingsSummary();
    }

    private void ApplyProcedurePaneLayout()
    {
        if (_isDraggingAnnotation) return;
        StepsPaneColumn.Width = new GridLength(_showStepsPane ? _stepsPaneWidth : 0);
        StepsResizeThumb.Visibility = _showStepsPane ? Visibility.Visible : Visibility.Collapsed;
        StepsPaneHost.Visibility = _showStepsPane ? Visibility.Visible : Visibility.Collapsed;
        InspectorPaneColumn.Width = new GridLength(_isInspectorPaneCollapsed ? 0 : 284);
        InspectorPaneHost.Visibility = _isInspectorPaneCollapsed ? Visibility.Collapsed : Visibility.Visible;
        StepsHeaderPanel.Visibility = DocumentTabsListView.Visibility = Visibility.Visible;
        DocumentTabsCompactListView.Visibility = Visibility.Collapsed;
        CollapseStepsButton.Visibility = AddStepButton.Visibility = Visibility.Visible;
        StepsPaneHintText.Visibility = Visibility.Collapsed;
        ToolbarStepsButton.Content = SafetyText("UiSteps");
        ToolbarOpacityValueText.Text = string.Format(CultureInfo.CurrentCulture, "{0:0} %", ToolbarOpacitySlider.Value * 100);
        QueueWorkspaceFit();
    }

    private void ApplyFloatingToolRailPosition(double left, double top)
    {
        // The rail is docked outside the image viewport; old floating positions remain compatible.
        FloatingToolRail.Margin = new Thickness(4);
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
                Width = 28,
                Height = 28,
                Padding = new Thickness(0),
                Content = string.Empty,
                FontSize = 9
            };
            button.Click += ShadeButton_Click;
            var tooltip = string.Format(CultureInfo.CurrentCulture, SafetyText("UiShade0"), i + 1);
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
        if (_isSavingSettings && section != ShellSection.Settings) return;
        _currentSection = section;
        CaptureSection.Visibility = section == ShellSection.Capture ? Visibility.Visible : Visibility.Collapsed;
        ProcedureSection.Visibility = section == ShellSection.Procedure ? Visibility.Visible : Visibility.Collapsed;
        SettingsSection.Visibility = section == ShellSection.Settings ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceFooterLeft.Visibility = section == ShellSection.Procedure ? Visibility.Visible : Visibility.Collapsed;
        foreach (var control in new UIElement[] { ToolbarZoomOutButton, ToolbarZoomResetButton, ToolbarZoomInButton, ToolbarFitButton, AnnotationCountText })
            control.Visibility = section == ShellSection.Procedure && !_showWelcome ? Visibility.Visible : Visibility.Collapsed;
        UpdateWelcomeSurface();
    }

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentSection == ShellSection.Settings)
        {
            ReturnToProcedureSection();
            return;
        }

        CaptureSettingsSnapshot();
        SetSection(ShellSection.Settings);
        SelectSettingsNavSection("General");
    }

    private void CaptureSettingsSnapshot()
    {
        _settingsDraft = App.Preferences.Clone();
        _settingsMetadata = _projectMetadata.Clone();
    }

    private void ReturnToProcedureSection()
    {
        if (_isSavingSettings) return;
        _settingsDraft = null;
        _settingsMetadata = null;
        var wasInitializing = _isInitializingUi;
        _isInitializingUi = true;
        try { ApplyPreferencesToUi(); ApplyChromeLabels(); SyncGuideMetadataEditors(); UpdateSettingsSummary(); }
        finally { _isInitializingUi = wasInitializing; }
        SetSection(ShellSection.Procedure);
    }

    private void SelectSettingsNavSection(string sectionKey)
    {
        _selectedSettingsSection = sectionKey;
        SetSettingsNavButtonState(SettingsGeneralNavButton, sectionKey == "General");
        SetSettingsNavButtonState(SettingsCaptureNavButton, sectionKey == "Capture");
        SetSettingsNavButtonState(SettingsExportNavButton, sectionKey == "Export");
        SetSettingsNavButtonState(SettingsGuideNavButton, sectionKey == "Guide");
        SetSettingsNavButtonState(SettingsPublicationNavButton, sectionKey == "Publication");
        SetSettingsNavButtonState(SettingsShortcutsNavButton, sectionKey == "Shortcuts");
        SetSettingsNavButtonState(SettingsAboutNavButton, sectionKey == "About");

        SetText(SettingsPageTitleText, sectionKey switch
        {
            "General" => SafetyText("UiGeneral"),
            "Capture" => SafetyText("UiCapture"),
            "Export" => SafetyText("UiExport"),
            "Guide" => SafetyText("GuideSettings"),
            "Publication" => SafetyText("UiPublication"),
            "Shortcuts" => SafetyText("UiShortcuts"),
            "About" => SafetyText("UiAbout"),
            _ => SafetyText("UiGeneral")
        });

        foreach (var (key, card) in new[] {
            ("General", SettingsGeneralCard), ("Capture", SettingsCaptureCard), ("Guide", SettingsGuideCard),
            ("Export", ExportSection), ("Publication", PublishSection), ("About", SettingsAboutCard) })
            card.Visibility = key == sectionKey ? Visibility.Visible : Visibility.Collapsed;
        SettingsScrollViewer.ChangeView(null, 0, null);
    }

    private void SetSettingsNavButtonState(Button button, bool isSelected)
    {
        button.Background = GetBrush(isSelected ? "ToolSelectedBrush" : "ShellPanelBrush");
        button.Foreground = GetBrush(isSelected ? "ToolSelectedTextBrush" : "ShellStrongTextBrush");
        button.BorderThickness = new Thickness(isSelected ? 3 : 1, 1, 1, 1);
        AutomationProperties.SetItemStatus(button, isSelected ? SafetyText("SelectedCategory") : string.Empty);
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
            StatusText.Text = SafetyText("UiHelpOpened");
            return;
        }

        SelectSettingsNavSection(sectionKey);
    }

    private async void SettingsSaveButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isSavingSettings || _settingsDraft is null || _settingsMetadata is null) return;
        _isSavingSettings = true;
        SettingsScrollViewer.IsEnabled = SettingsSaveButton.IsEnabled = SettingsCancelButton.IsEnabled = false;
        try
        {
            var preferences = _settingsDraft.Clone();
            await Task.Run(preferences.Save);
            App.Preferences.CopyFrom(preferences);
            if (ProjectJson.Serialize(_projectMetadata) != ProjectJson.Serialize(_settingsMetadata))
            {
                PushUndoState(CaptureProjectState());
                _projectMetadata = _settingsMetadata.Clone();
                RefreshProjectDirtyIndicators();
            }
            _isSavingSettings = false;
            ReturnToProcedureSection();
            StatusText.Text = SafetyText("UiSettingsSaved");
        }
        catch (Exception error) { AppDiagnostics.Record(error); StatusText.Text = SafetyText("OperationFailed"); }
        finally { _isSavingSettings = false; SettingsScrollViewer.IsEnabled = SettingsSaveButton.IsEnabled = SettingsCancelButton.IsEnabled = true; }
    }

    private void SettingsCancelButton_Click(object sender, RoutedEventArgs e)
    {
        ReturnToProcedureSection();
        StatusText.Text = SafetyText("UiSettingsCancelled");
    }

    private void FloatingToolRailDragThumb_DragDelta(object sender, DragDeltaEventArgs e)
    {
        var currentLeft = EditingPreferences.FloatingToolRailLeft;
        var currentTop = EditingPreferences.FloatingToolRailTop;
        ApplyFloatingToolRailPosition(currentLeft + e.HorizontalChange, currentTop + e.VerticalChange);
    }

    private void FloatingToolRailDragThumb_DragCompleted(object sender, DragCompletedEventArgs e)
    {
        ApplyFloatingToolRailPosition(EditingPreferences.FloatingToolRailLeft, EditingPreferences.FloatingToolRailTop);
        SavePreferencesIfNeeded();
    }

    private void ShowHelpSectionMenuItem_Click(object sender, RoutedEventArgs e)
    {
        SetSection(ShellSection.Capture);
        HelpShortcutsTitleText.StartBringIntoView();
        StatusText.Text = SafetyText("UiHelpOpened");
    }

    private void ReturnToProcedureMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (_currentSection == ShellSection.Settings)
        {
            ReturnToProcedureSection();
        }
        else
        {
            SetSection(ShellSection.Procedure);
        }

        StatusText.Text = SafetyText("UiBackToTheProcedure");
    }

    private void CaptureShortcutMenuItem_Click(object sender, RoutedEventArgs e)
    {
        SetSection(ShellSection.Capture);
        HelpShortcutsTitleText.StartBringIntoView();

        var description = sender is MenuFlyoutItem { Tag: string tag }
            ? tag switch
            {
                "PrintScreen" => SafetyText("UiPrintScreenCopiesTheWholeScreenTo"),
                "AltPrintScreen" => SafetyText("UiAltPrintScreenCopiesTheActiveWindow"),
                "WinPrintScreen" => SafetyText("UiWinPrintScreenSavesTheCaptureTo"),
                "WinShiftS" => SafetyText("UiWinShiftSOpensSnippingTool"),
                "FnWinSpace" => SafetyText("UiFnWinSpaceAlternateOnSomeKeyboards"),
                _ => SafetyText("UiScreenshotHelpOpened")
            }
            : SafetyText("UiScreenshotHelpOpened");

        StatusText.Text = description;
    }

    private void CollapseStepsButton_Click(object sender, RoutedEventArgs e)
    {
        _showStepsPane = !_showStepsPane;
        ApplyProcedurePaneLayout();
    }

    private void InspectorToggleButton_Click(object sender, RoutedEventArgs e)
    {
        _isInspectorPaneCollapsed = !_isInspectorPaneCollapsed;
        _inspectorOpenedManually = !_isInspectorPaneCollapsed;
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

    private void ZoomInButton_Click(object sender, RoutedEventArgs e) => SetManualZoom(_currentZoom + 0.1);

    private void ZoomOutButton_Click(object sender, RoutedEventArgs e) => SetManualZoom(_currentZoom - 0.1);

    private void ZoomResetButton_Click(object sender, RoutedEventArgs e) => SetManualZoom(1.0);

    private void FitToViewportButton_Click(object sender, RoutedEventArgs e)
    {
        _autoFitPreview = true;
        FitPreviewToViewport();
    }

    private void FitPreviewToViewport()
    {
        var contentWidth = _currentDocument is null ? PreviewViewport.Width : GetSceneWidth(_currentDocument) + FloatingToolRailGutterWidth;
        var contentHeight = _currentDocument is null ? PreviewViewport.Height : GetSceneHeight(_currentDocument);
        var availableWidth = Math.Max(1, CanvasScrollViewer.ActualWidth - 24);
        var availableHeight = Math.Max(1, CanvasScrollViewer.ActualHeight - 24);

        if (contentWidth <= 0 || contentHeight <= 0 || availableWidth <= 0 || availableHeight <= 0)
        {
            SetZoom(1.0);
            return;
        }

        var fitZoom = Math.Min(availableWidth / contentWidth, availableHeight / contentHeight);
        SetZoom(Math.Min(1, fitZoom));
    }

    private void SetZoom(double zoom)
    {
        _currentZoom = Math.Clamp(zoom, 0.05, 2.0);
        SceneScaleTransform.ScaleX = _currentZoom;
        SceneScaleTransform.ScaleY = _currentZoom;
        if (_currentDocument is not null) ApplyCropState(_currentDocument);
        ZoomLevelText.Text = string.Format(CultureInfo.CurrentCulture, "{0:0} %", _currentZoom * 100);
        UpdateWorkspaceZoomLabel();
        UpdateStatus();
    }

    private void DocumentTitleTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        if (EditingMetadata.Title == textBox.Text) return;
        EditingMetadata.Title = textBox.Text;
        GuideMetadataChanged();
    }

    private void StepNoteTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        if (_currentDocument.StepNote == textBox.Text) return;
        _currentDocument.StepNote = textBox.Text;
        if (_currentDocument.ImageBytes is null) RequestScenePreview();
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
            if (EditingMetadata.TemplateType == type) return;
            if (_settingsDraft is null) PushUndoState(CaptureProjectState());
            EditingMetadata.TemplateType = type;
            GuideMetadataChanged();
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
            if (EditingMetadata.TemplateType == type) return;
            if (_settingsDraft is null) PushUndoState(CaptureProjectState());
            EditingMetadata.TemplateType = type;
            GuideMetadataChanged();
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

        if (EditingMetadata.Audience == textBox.Text) return;
        EditingMetadata.Audience = textBox.Text;
        GuideMetadataChanged();
    }

    private void AuthorTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        if (EditingMetadata.Author == textBox.Text) return;
        EditingMetadata.Author = textBox.Text;
        GuideMetadataChanged();
    }

    private void DocumentVersionTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null || sender is not TextBox textBox)
        {
            return;
        }

        if (EditingMetadata.DocumentVersion == textBox.Text) return;
        EditingMetadata.DocumentVersion = textBox.Text;
        GuideMetadataChanged();
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

        EditingPreferences.ExportFolder = textBox.Text;
        SavePreferencesIfNeeded();

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
            StatusText.Text = SafetyText("UiFolderSelectionCancelled");
            return;
        }

        EditingPreferences.ExportFolder = folder.Path;
        SavePreferencesIfNeeded();
        if (ExportFolderTextBox is not null)
        {
            ExportFolderTextBox.Text = folder.Path;
        }

        UpdateSettingsSummary();
        StatusText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("UiExportFolder0"), folder.Path);
    }

    private void ExportFormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || ExportFormatComboBox.SelectedItem is not ComboBoxItem { Tag: ExportFormatChoice format })
        {
            return;
        }

        EditingPreferences.DefaultExportFormat = format;
        SavePreferencesIfNeeded();
        UpdateSettingsSummary();
    }

    private void ThemeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || ThemeComboBox.SelectedItem is not ComboBoxItem { Tag: ThemeChoice theme })
        {
            return;
        }

        EditingPreferences.Theme = theme;
        SavePreferencesIfNeeded();
        RootGrid.RequestedTheme = EditingPreferences.GetRequestedTheme();
        UpdateSettingsSummary();
    }

    private void LanguageComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializingUi || LanguageComboBox.SelectedItem is not ComboBoxItem { Tag: LanguageChoice language })
        {
            return;
        }

        EditingPreferences.Language = language;
        SavePreferencesIfNeeded();
        ApplyChromeLabels();
        UpdateSettingsSummary();
    }

    private void AutoImportClipboardToggleSwitch_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isInitializingUi)
        {
            return;
        }

        EditingPreferences.AutoImportClipboard = AutoImportClipboardToggleSwitch.IsOn;
        SavePreferencesIfNeeded();
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
            _selectedAnnotation.Opacity = _selectedAnnotation.Kind == AnnotationKind.Mask ? 1 : SelectedOpacitySlider.Value;
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

        var moveUp = new MenuFlyoutItem { Text = SafetyText("UiMoveUp") };
        moveUp.Click += (_, _) => MoveDocument(document, -1);
        flyout.Items.Add(moveUp);

        var moveDown = new MenuFlyoutItem { Text = SafetyText("UiMoveDown") };
        moveDown.Click += (_, _) => MoveDocument(document, 1);
        flyout.Items.Add(moveDown);

        var duplicate = new MenuFlyoutItem { Text = SafetyText("UiDuplicate") };
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

        var delete = new MenuFlyoutItem { Text = SafetyText("UiDelete") };
        delete.Click += async (_, _) => await TryCloseDocumentAsync(document);
        flyout.Items.Add(delete);

        flyout.ShowAt((FrameworkElement)sender);
    }

    private void AddLegendButton_Click(object sender, RoutedEventArgs e)
    {
        SelectTool(EditorTool.Sticker);
        StatusText.Text = SafetyText("UiClickTheCanvasToAddASticker");
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
                Title = SafetyText("UiDeleteTheSticker"),
                Content = SafetyText("UiDoYouWantToRenumberTheRemaining"),
                PrimaryButtonText = SafetyText("UiRenumber"),
                SecondaryButtonText = SafetyText("UiDeleteWithoutRenumbering"),
                CloseButtonText = SafetyText("UiCancel"),
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
            if (_settingsDraft is not null || e.OriginalSource is not TextBox)
            {
                ReturnToProcedureSection();
                StatusText.Text = SafetyText("UiBackToTheProcedure");
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
        StatusText.Text = SafetyText("UiTextEditableInThePreview");
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
        SyncGuideMetadataEditors();
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
        AnnotationTextBox.Text = document.AnnotationText;
        StickerModeComboBox.SelectedIndex = document.StickerModeIndex == 1 ? 1 : 0;
        OpacitySlider.Value = document.DefaultOpacity;
        ToolbarOpacitySlider.Value = document.DefaultOpacity;
        ToolbarOpacityValueText.Text = string.Format(CultureInfo.CurrentCulture, "{0:0} %", document.DefaultOpacity * 100);
        SelectedOpacitySlider.Value = _selectedAnnotation?.Opacity ?? document.DefaultOpacity;
        ResetStickerNumberOnColorChangeCheckBox.IsChecked = document.ResetStickerNumberOnColorChange;

        ExportFolderTextBox.Text = EditingPreferences.ExportFolder;
        ExportFormatComboBox.SelectedIndex = (int)EditingPreferences.DefaultExportFormat;
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
                    ? SafetyText("UiToFillIn")
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

        var selectionChanged = !ReferenceEquals(_selectedAnnotation, annotation);
        _selectedAnnotation = annotation;
        if (selectionChanged && !_inspectorOpenedManually) _isInspectorPaneCollapsed = annotation?.Kind is not (AnnotationKind.Text or AnnotationKind.Sticker);
        ApplyProcedurePaneLayout();
        _isSyncingObjectList = true;
        AnnotationObjectsList.SelectedItem = annotation;
        _isSyncingObjectList = false;
        if (_currentDocument is null)
        {
            return;
        }

        _isSyncingAnnotationSelectionUi = true;
        try
        {
            if (annotation is null)
            {
                LegendPanel.Visibility = Visibility.Visible;
                StepDetailsPanel.Visibility = Visibility.Visible;
                AnnotationContextPanel.Visibility = Visibility.Collapsed;
                SelectionOutlineHalo.Visibility = SelectionOutline.Visibility = Visibility.Collapsed;
                StepDetailsSubtitleText.Text = SafetyText("UiStepTitleAndNote");
                SelectedAnnotationSummaryText.Text = SafetyText("UiNoObjectSelected");
                SelectedColorLabelText.Visibility = Visibility.Collapsed;
                SelectedColorText.Visibility = Visibility.Collapsed;
                SelectedOpacitySlider.Value = _currentDocument.DefaultOpacity;
                return;
            }

            StepDetailsPanel.Visibility = Visibility.Collapsed;
            AnnotationContextPanel.Visibility = Visibility.Visible;
            StepDetailsSubtitleText.Text = SafetyText("UiSelectedObject");
            AnnotationPanelTitleText.Text = annotation.Kind switch
            {
                AnnotationKind.Sticker => SafetyText("UiSticker"),
                AnnotationKind.Text => SafetyText("UiText_0420e350"),
                AnnotationKind.Rectangle => SafetyText("UiRectangle"),
                AnnotationKind.Ellipse => SafetyText("UiEllipse"),
                AnnotationKind.ArrowStraight => SafetyText("UiStraightArrow"),
                AnnotationKind.ArrowCurved => SafetyText("UiCurvedArrow"),
                AnnotationKind.Focus => SafetyText("UiFocus"),
                AnnotationKind.Mask => SafetyText("UiMask"),
                _ => SafetyText("UiObject")
            };

            var isSticker = annotation.Kind == AnnotationKind.Sticker;
            SelectedStickerValueHost.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            LegendPanel.Visibility = isSticker ? Visibility.Visible : Visibility.Collapsed;
            var isText = annotation.Kind == AnnotationKind.Text;
            var isStrokeBased = annotation.Kind is AnnotationKind.Rectangle or AnnotationKind.Ellipse or AnnotationKind.ArrowStraight or AnnotationKind.ArrowCurved or AnnotationKind.Focus or AnnotationKind.Mask;
            var paletteName = _palettes.TryGetValue(annotation.PaletteKey, out var selectedPalette)
                ? selectedPalette.DisplayName
                : annotation.PaletteKey;

            SelectedColorLabelText.Visibility = Visibility.Collapsed;
            SelectedColorText.Visibility = Visibility.Collapsed;
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
            StrokeSummaryText.Visibility = Visibility.Collapsed;
            CropStateText.Visibility = (_currentDocument.PendingCropRect is not null || _currentDocument.AppliedCropRect is not null) ? Visibility.Visible : Visibility.Collapsed;
            ApplyCropButton.Visibility = _currentDocument.PendingCropRect is not null ? Visibility.Visible : Visibility.Collapsed;
            ResetCropButton.Visibility = (_currentDocument.PendingCropRect is not null || _currentDocument.AppliedCropRect is not null) ? Visibility.Visible : Visibility.Collapsed;

            StickerValueLabelText.Text = SafetyText("UiStickerValue");
            StickerValueText.Text = annotation.Text;
            StickerLegendLabelText.Text = SafetyText("UiStickerLegend");
            StickerLegendTextBox.Text = annotation.LegendText;
            StickerLegendTextBox.PlaceholderText = SafetyText("UiDescribeWhatThisStickerPointsTo");
            AnnotationTextLabelText.Text = SafetyText("UiText_0420e350");
            AnnotationTextBox.Text = annotation.Text;
            AnnotationTextBox.PlaceholderText = SafetyText("UiEnterText");
            SelectedColorLabelText.Text = SafetyText("UiColor");
            SelectedColorText.Text = string.Format(CultureInfo.CurrentCulture, "{0} · {1}", paletteName, annotation.PaletteShadeIndex + 1);
            SelectedAnnotationSummaryText.Text = annotation.Kind switch
            {
                AnnotationKind.Sticker or AnnotationKind.Text => $"{annotation.Text} · {annotation.Opacity:P0}",
                AnnotationKind.Rectangle or AnnotationKind.Ellipse => $"{annotation.Bounds.Width:0} x {annotation.Bounds.Height:0} · {annotation.Opacity:P0}",
                _ => annotation.Opacity.ToString("P0", CultureInfo.CurrentCulture)
            };

            SelectedOpacityLabelText.Visibility = SelectedOpacitySlider.Visibility = annotation.Kind == AnnotationKind.Mask ? Visibility.Collapsed : Visibility.Visible;
            SelectedOpacitySlider.Value = annotation.Opacity;
            StrokeThicknessSlider.Value = annotation.StrokeThickness;
            AnnotationTextBox.Text = isText ? annotation.Text : string.Empty;
            UpdateSelectionOutline(annotation);
        }
        finally
        {
            _isSyncingAnnotationSelectionUi = false;
            UpdateWorkspaceContext();
        }
    }

    private void UpdateSelectionOutline(AnnotationModel annotation)
    {
        var bounds = GetAnnotationVisualBounds(annotation, _draggingAnnotationElement);

        if (bounds.Width <= 0 || bounds.Height <= 0)
        {
            SelectionOutlineHalo.Visibility = SelectionOutline.Visibility = Visibility.Collapsed;
            return;
        }

        foreach (var outline in new[] { SelectionOutlineHalo, SelectionOutline })
        {
            outline.Visibility = Visibility.Visible;
            outline.Width = bounds.Width + 12;
            outline.Height = bounds.Height + 12;
            Canvas.SetLeft(outline, Math.Max(0, bounds.X - 6));
            Canvas.SetTop(outline, Math.Max(0, bounds.Y - 6));
        }
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

            BeginAnnotationDrag(element, annotation, e);
            SelectAnnotation(annotation);
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

        _annotationMoveSnapshot = CaptureProjectState();
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

        RequestScenePreview();
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

        if (_annotationMoveSnapshot is { } before && !AreProjectStatesEquivalent(before, CaptureProjectState())) PushUndoState(before);
        _annotationMoveSnapshot = null;
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
            BaseTitle = string.Format(CultureInfo.CurrentCulture, SafetyText("CopyTitle"), source.BaseTitle),
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
        CommitPendingTextEdits();
        try
        {
            var rootFolder = EditingPreferences.ExportFolder;
            if (string.IsNullOrWhiteSpace(rootFolder))
                rootFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SnapSlate-Exports");
            var package = await BuildGuidePackageAsync(rootFolder, formats);
            if (package is not null)
                StatusText.Text = string.Format(CultureInfo.CurrentCulture,
                    SafetyText("UiExportComplete0In1"),
                    string.Join(", ", package.GeneratedFormats), package.PackageFolder);
        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            StatusText.Text = SafetyText("ExportFailed");
        }
    }

    private string GetTemplateLabel(GuideTemplateType templateType)
    {
        return templateType switch
        {
            GuideTemplateType.InternalProcedure => SafetyText("UiInternalProcedure"),
            GuideTemplateType.SupportPlaybook => SafetyText("UiSupportPlaybook"),
            GuideTemplateType.ReleaseNotes => SafetyText("UiReleaseNotes"),
            _ => SafetyText("UiUserManual")
        };
    }

    private static string EscapeHtml(string value) => GuideTextExporter.EscapeHtml(value);

    private string BuildMarkdown(GuideExportManifest manifest, bool includeImageReferences = true) =>
        GuideTextExporter.BuildMarkdown(manifest, includeImageReferences, UseEnglish);

    private string BuildHtml(GuideExportManifest manifest) => GuideTextExporter.BuildHtml(manifest, UseEnglish);

}
