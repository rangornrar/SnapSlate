using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Automation.Peers;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Windows.System;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private readonly Microsoft.UI.System.ThemeSettings _windowThemeSettings;
    private readonly System.Collections.Generic.HashSet<Border> _stepCards = [];
    private string _selectedSettingsSection = "General";
    private double _stepsPaneWidth = 320;
    private bool _showWelcome = true;
    private bool _toolOptionsOpen;
    private EditorTool _lastArrowTool = EditorTool.ArrowStraight;
    private EditorTool _lastShapeTool = EditorTool.Rectangle;

    private void InitializeWorkflowTheme()
    {
        RootGrid.ActualThemeChanged += WorkflowThemeChanged;
        _windowThemeSettings.Changed += WorkflowSystemThemeChanged;
        RefreshWorkflowTheme();
    }

    private void WorkflowThemeChanged(FrameworkElement sender, object args) => QueueWorkflowThemeRefresh();

    private void WorkflowSystemThemeChanged(Microsoft.UI.System.ThemeSettings sender, object args) => QueueWorkflowThemeRefresh();

    private void QueueWorkflowThemeRefresh() => DispatcherQueue.TryEnqueue(() =>
    {
        if (!_editorClosed) RefreshWorkflowTheme();
    });

    private void RefreshWorkflowTheme()
    {
        var wasInitializing = _isInitializingUi;
        _isInitializingUi = true;
        try
        {
            ApplyChromeLabels();
            SelectTool(_currentTool);
            RefreshShadeStrip();
            AppWindow.TitleBar.ButtonForegroundColor = GetBrush("ShellStrongTextBrush").Color;
            AppWindow.TitleBar.ButtonInactiveForegroundColor = GetBrush("ShellMutedTextBrush").Color;
            AppWindow.TitleBar.ButtonHoverForegroundColor = GetBrush("ShellStrongTextBrush").Color;
            AppWindow.TitleBar.ButtonHoverBackgroundColor = GetBrush("ShellPanelAltBrush").Color;
        }
        finally { _isInitializingUi = wasInitializing; }
    }

    private void LocalizePalettes()
    {
        var wasSyncing = _isSyncingPaletteUi;
        _isSyncingPaletteUi = true;
        try
        {
            PopulatePaletteOrder();
            if (_currentDocument is not null)
                ToolbarPaletteComboBox.SelectedItem = _paletteOrder.FirstOrDefault(p => p.Key == _currentDocument.SelectedPaletteKey);
            BuildPaletteStrip();
        }
        finally { _isSyncingPaletteUi = wasSyncing; }
    }

    private void BuildQuickColors()
    {
        QuickColorStrip.Children.Clear();
        foreach (var key in new[] { "Sunset", "Ember", "Mint", "Sky", "Grape" })
        {
            var button = new Button
            {
                Tag = key, Width = 30, Height = 30, Padding = new Thickness(0),
                Background = new Microsoft.UI.Xaml.Media.SolidColorBrush(_palettes[key].Shades[3])
            };
            SetButtonTooltip(button, SafetyText("Palette" + key));
            AutomationProperties.SetAutomationId(button, "QuickColor" + key);
            button.Click += async (_, _) => await ApplyPaletteSelectionAsync(key, 3);
            QuickColorStrip.Children.Add(button);
        }
        RefreshQuickColors();
    }

    private void RefreshQuickColors()
    {
        foreach (var button in QuickColorStrip.Children.OfType<Button>())
        {
            var selected = (string)button.Tag == _currentDocument?.SelectedPaletteKey && _currentDocument.SelectedPaletteShadeIndex == 3;
            button.BorderThickness = new Thickness(selected ? 3 : 1);
            button.BorderBrush = GetBrush(selected ? "ToolSelectedBrush" : "ShellStrokeBrush");
            AutomationProperties.SetItemStatus(button, selected ? SafetyText("SelectedColor") : string.Empty);
        }
    }

    private void ApplyWorkflowLabels()
    {
        BuildQuickColors();
        MoreColorsButton.Content = SafetyText("MoreColors");
        SetButtonTooltip(MoreColorsButton, SafetyText("MoreColorsHelp"));
        WelcomeTitleText.Text = SafetyText("WelcomeTitle");
        WelcomeHelpText.Text = SafetyText("WelcomeHelp");
        WelcomePasteButton.Content = SafetyText("WelcomePaste");
        WelcomeImageButton.Content = SafetyText("WelcomeImage");
        WelcomeProjectButton.Content = SafetyText("WelcomeProject");
        WelcomeExampleButton.Content = SafetyText("WelcomeExample");
        SettingsGuideNavText.Text = SafetyText("GuideNav");
        AutomationProperties.SetName(SettingsGuideNavButton, SafetyText("GuideSettings"));
        AutomationProperties.SetName(StepsResizeThumb, SafetyText("ResizeSteps"));
        ToolTipService.SetToolTip(StepsResizeThumb, SafetyText("ResizeSteps"));
        foreach (var item in new[] { ToolVariantArrowStraight, ToolVariantArrowCurved, ToolVariantRectangle, ToolVariantEllipse })
            item.Text = SafetyText(_toolLabels[Enum.Parse<EditorTool>((string)item.Tag)]);
        AutomationProperties.SetName(ArrowVariantsButton, SafetyText("ArrowVariants"));
        AutomationProperties.SetName(ShapeVariantsButton, SafetyText("ShapeVariants"));
        AutomationProperties.SetName(SettingsCloseButton, SafetyText("UiCancel"));
        foreach (var (button, label) in new[] {
            (SettingsGeneralNavButton, SettingsGeneralNavText), (SettingsCaptureNavButton, SettingsCaptureNavText),
            (SettingsExportNavButton, SettingsExportNavText), (SettingsPublicationNavButton, SettingsPublicationNavText),
            (SettingsShortcutsNavButton, SettingsShortcutsNavText), (SettingsAboutNavButton, SettingsAboutNavText) })
            AutomationProperties.SetName(button, label.Text);
        if (_currentSection == ShellSection.Settings) SelectSettingsNavSection(_selectedSettingsSection);
        UpdateGroupedTools();
        UpdateWelcomeSurface();
    }

    private void UpdateWelcomeSurface()
    {
        UpdateStepSelectionChrome();
        WelcomeHost.Visibility = _showWelcome ? Visibility.Visible : Visibility.Collapsed;
        EditorWorkspaceGrid.Visibility = _showWelcome ? Visibility.Collapsed : Visibility.Visible;
        CopyImageButton.IsEnabled = CopyImageMenuItem.IsEnabled = !_showWelcome && !_isCopyingImage;
        ExportCurrentButton.IsEnabled = !_showWelcome;
        foreach (var control in new UIElement[] { ToolbarZoomOutButton, ToolbarZoomResetButton, ToolbarZoomInButton, ToolbarFitButton, AnnotationCountText })
            control.Visibility = !_showWelcome && _currentSection == ShellSection.Procedure ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceFooterLeft.Visibility = _showWelcome || _currentSection != ShellSection.Procedure ? Visibility.Collapsed : Visibility.Visible;
        StepPositionText.Text = _showWelcome || _currentDocument is null ? string.Empty : string.Format(
            CultureInfo.CurrentCulture, SafetyText("StepPosition"), _documents.IndexOf(_currentDocument) + 1, _documents.Count);
    }

    private void StepCard_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is Border card) _stepCards.Add(card);
        UpdateStepSelectionChrome();
    }

    private void StepCard_Unloaded(object sender, RoutedEventArgs e)
    {
        if (sender is Border card) _stepCards.Remove(card);
    }

    private void UpdateStepSelectionChrome()
    {
        foreach (var card in _stepCards)
        {
            var selected = ReferenceEquals(card.DataContext, _currentDocument);
            card.BorderBrush = GetBrush(selected ? "ToolSelectedBrush" : "ShellStrokeBrush");
            card.BorderThickness = new Thickness(selected ? 4 : 1, 1, 1, 1);
        }
    }

    private void StepsResizeThumb_DragDelta(object sender, DragDeltaEventArgs e) => ResizeStepsPane(e.HorizontalChange);

    private void StepsResizeThumb_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key is VirtualKey.Left or VirtualKey.Right)
        {
            ResizeStepsPane(e.Key == VirtualKey.Left ? -16 : 16);
            e.Handled = true;
        }
    }

    private void ResizeStepsPane(double delta)
    {
        _stepsPaneWidth = Math.Clamp(_stepsPaneWidth + delta, 240, 420);
        ApplyProcedurePaneLayout();
    }

    private void UpdateToolButtonState(Button button, EditorTool tool)
    {
        var selected = _currentTool == tool;
        _iconBrushOverride = GetBrush(selected ? "ToolSelectedTextBrush" : "ShellStrongTextBrush");
        button.Content = CreateToolIcon(tool);
        _iconBrushOverride = null;
        button.Background = GetBrush(selected ? "ToolSelectedBrush" : "ShellPanelBrush");
        button.BorderBrush = GetBrush(selected ? "ToolSelectedBrush" : "ShellStrokeBrush");
        button.BorderThickness = new Thickness(selected ? 2 : 1);
    }

    private void UpdateGroupedTools()
    {
        void Update(Button button, EditorTool tool)
        {
            button.Tag = tool.ToString();
            UpdateToolButtonState(button, tool);
            SetButtonTooltip(button, SafetyText(_toolLabels[tool]) + (tool is EditorTool.ArrowStraight or EditorTool.ArrowCurved ? " (A)" : tool == EditorTool.Rectangle ? " (R)" : " (O)") + " · " + SafetyText("ToolVariantsHelp"));
        }
        Update(ToolbarArrowStraightButton, _lastArrowTool);
        Update(ToolbarRectangleButton, _lastShapeTool);
    }

    private void ShowToolVariants_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button) return;
        var toolButton = ReferenceEquals(button, ArrowVariantsButton) ? ToolbarArrowStraightButton : ToolbarRectangleButton;
        toolButton.ContextFlyout.ShowAt(button);
    }

    private void GroupedTool_Click(object sender, RoutedEventArgs e)
    {
        // A normal click activates the remembered variant; the menu offers the alternatives.
        if (sender is Button { Tag: string tag }) SelectTool(Enum.Parse<EditorTool>(tag));
    }

    private void ToolVariant_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuFlyoutItem { Tag: string tag }) SelectTool(Enum.Parse<EditorTool>(tag));
    }

    private void ToolOptions_Opened(object sender, object e) => _toolOptionsOpen = true;

    private void ToolOptions_Closed(object sender, object e) => _toolOptionsOpen = false;

    private void InitializeToolShortcuts()
    {
        foreach (var (key, tool) in new[] {
            (VirtualKey.V, EditorTool.Select), (VirtualKey.T, EditorTool.Text),
            (VirtualKey.A, EditorTool.ArrowStraight), (VirtualKey.R, EditorTool.Rectangle),
            (VirtualKey.O, EditorTool.Ellipse), (VirtualKey.M, EditorTool.Mask) })
        {
            AddKeyboardAccelerator(key, VirtualKeyModifiers.None, (_, args) =>
            {
                if (_toolOptionsOpen || IsTextInputFocused() || _showWelcome || _currentSection != ShellSection.Procedure || _isProjectDialogOpen) return;
                SelectTool(tool == EditorTool.ArrowStraight ? _lastArrowTool : tool);
                args.Handled = true;
            });
        }
    }

    private async void OpenExportDialog_Click(object sender, RoutedEventArgs e)
    {
        if (_showWelcome || _currentDocument is null || _isProjectDialogOpen || RootGrid.XamlRoot is null) return;
        CommitPendingTextEdits();
        var scope = new ComboBox { Header = SafetyText("ExportScope"), HorizontalAlignment = HorizontalAlignment.Stretch };
        scope.Items.Add(SafetyText("ExportActiveCapture"));
        scope.Items.Add(string.Format(CultureInfo.CurrentCulture, SafetyText("ExportWholeGuide"), _documents.Count));
        scope.SelectedIndex = 0;
        var format = new ComboBox { Header = SafetyText("ExportFormatLabel"), HorizontalAlignment = HorizontalAlignment.Stretch };
        var choices = new[] { ExportFormatChoice.Png, ExportFormatChoice.Markdown, ExportFormatChoice.Html, ExportFormatChoice.Pdf, ExportFormatChoice.Docx };
        foreach (var choice in choices) format.Items.Add(GetExportFormatDisplayName(choice));
        format.Items.Add(SafetyText("ExportGuideFormats"));
        format.SelectedIndex = Math.Max(0, Array.IndexOf(choices, App.Preferences.DefaultExportFormat));
        var folder = new TextBox { Header = SafetyText("ExportDestination"), Text = App.Preferences.ExportFolder ?? string.Empty };
        if (string.IsNullOrWhiteSpace(folder.Text))
            folder.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SnapSlate-Exports");
        var browse = new Button { Content = SafetyText("UiBrowse") };
        AutomationProperties.SetName(scope, SafetyText("ExportScope"));
        AutomationProperties.SetName(format, SafetyText("ExportFormatLabel"));
        AutomationProperties.SetName(folder, SafetyText("ExportDestination"));
        AutomationProperties.SetAutomationId(scope, "ExportScope");
        AutomationProperties.SetAutomationId(format, "ExportFormat");
        AutomationProperties.SetAutomationId(folder, "ExportDestination");
        browse.Click += async (_, _) =>
        {
            try
            {
                var picker = new FolderPicker();
                picker.FileTypeFilter.Add("*");
                InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
                var selected = await picker.PickSingleFolderAsync();
                if (selected is not null) folder.Text = selected.Path;
            }
            catch (Exception error) { AppDiagnostics.Record(error); }
        };
        var errorText = new TextBlock { TextWrapping = TextWrapping.Wrap, Visibility = Visibility.Collapsed };
        AutomationProperties.SetLiveSetting(errorText, AutomationLiveSetting.Polite);
        var content = new StackPanel { Spacing = 14, MinWidth = 360 };
        content.Children.Add(scope);
        content.Children.Add(format);
        content.Children.Add(folder);
        content.Children.Add(browse);
        content.Children.Add(new TextBlock { Text = SafetyText("ExportSubfolderHelp"), TextWrapping = TextWrapping.Wrap });
        content.Children.Add(errorText);
        var dialog = new ContentDialog
        {
            XamlRoot = RootGrid.XamlRoot, Title = SafetyText("ExportAction"), Content = content,
            PrimaryButtonText = SafetyText("ExportAction"), CloseButtonText = SafetyText("UiCancel"),
            DefaultButton = ContentDialogButton.Primary, RequestedTheme = RootGrid.ActualTheme
        };
        dialog.PrimaryButtonClick += (_, args) =>
        {
            if (Path.IsPathFullyQualified(folder.Text.Trim())) return;
            args.Cancel = true;
            errorText.Text = SafetyText("ExportFolderRequired");
            errorText.Visibility = Visibility.Visible;
            folder.Focus(FocusState.Programmatic);
        };
        _isProjectDialogOpen = true;
        ContentDialogResult result;
        try { result = await dialog.ShowAsync(); }
        finally { _isProjectDialogOpen = false; }
        if (result != ContentDialogResult.Primary) return;
        var documents = scope.SelectedIndex == 0 ? new[] { _currentDocument } : _documents.ToArray();
        var formats = format.SelectedIndex == choices.Length ? choices.Skip(1).ToArray() : new[] { choices[format.SelectedIndex] };
        try
        {
            var package = await BuildGuidePackageAsync(folder.Text.Trim(), formats, documents);
            if (package is not null) StatusText.Text = string.Format(CultureInfo.CurrentCulture,
                SafetyText("UiExportComplete0In1"), string.Join(", ", package.GeneratedFormats), package.PackageFolder);
        }
        catch (Exception error) { AppDiagnostics.Record(error); StatusText.Text = SafetyText("ExportFailed"); }
    }
}
