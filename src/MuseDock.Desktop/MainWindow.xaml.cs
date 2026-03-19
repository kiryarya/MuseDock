using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using MuseDock.Desktop.Models;
using MuseDock.Desktop.Services;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;

namespace MuseDock.Desktop;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private const int WmMouseHWheel = 0x020E;
    private const double HorizontalScrollStep = 72d;
    private readonly AssetLibraryService _libraryService = new();
    private readonly LibraryMetadataStore _metadataStore = new();
    private readonly ObservableCollection<NavigationItem> _drives = [];
    private readonly ObservableCollection<DirectoryTreeNode> _directoryRoots = [];
    private readonly ObservableCollection<ExplorerTabState> _tabs = [];
    private readonly WorkspacePaneState[] _workspacePanes;
    private readonly ObservableCollection<string> _activityLog = [];
    private ExplorerTabState? _selectedTab;
    private WorkspacePaneState? _selectedWorkspace;
    private DirectoryTreeNode? _selectedDirectoryNode;
    private PendingTransferOperation? _pendingTransfer;
    private bool _suppressTreeSelection;
    private Point? _tabDragStartPoint;
    private ExplorerTabState? _dragTabCandidate;

    public MainWindow()
    {
        _workspacePanes =
        [
            new WorkspacePaneState(0) { IsVisible = true },
            new WorkspacePaneState(1),
            new WorkspacePaneState(2),
            new WorkspacePaneState(3)
        ];

        foreach (var pane in _workspacePanes)
        {
            pane.PropertyChanged += WorkspacePane_PropertyChanged;
        }

        _selectedWorkspace = _workspacePanes[0];
        SourceInitialized += MainWindow_SourceInitialized;
        InitializeComponent();
        DataContext = this;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<ExplorerTabState> Tabs => _tabs;

    public ObservableCollection<DirectoryTreeNode> DirectoryRoots => _directoryRoots;

    public ObservableCollection<string> ActivityLog => _activityLog;

    public WorkspacePaneState Workspace0 => _workspacePanes[0];

    public WorkspacePaneState Workspace1 => _workspacePanes[1];

    public WorkspacePaneState Workspace2 => _workspacePanes[2];

    public WorkspacePaneState Workspace3 => _workspacePanes[3];

    public ExplorerTabState? SelectedTab
    {
        get => _selectedTab;
        set
        {
            if (_selectedTab == value)
            {
                return;
            }

            _selectedTab = value;
            OnPropertyChanged();

            if (_selectedTab is not null)
            {
                var owningPane = FindWorkspaceForTab(_selectedTab);
                if (owningPane is not null)
                {
                    _selectedWorkspace = owningPane;
                    if (!ReferenceEquals(owningPane.SelectedTab, _selectedTab))
                    {
                        owningPane.SelectedTab = _selectedTab;
                    }
                }

                SelectDirectoryNodeForPath(_selectedTab.CurrentLocationPath);
            }
        }
    }

    public bool HasPendingTransfer => _pendingTransfer is not null;

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        UpdateWorkspaceLayout();
        await InitializeExplorerAsync();
    }

    private async Task InitializeExplorerAsync()
    {
        try
        {
            LoadDrives();
            LoadDirectoryRoots();
            AddLog("ドライブ一覧を読み込みました。");

            var initialPath = _drives.FirstOrDefault()?.Path;
            if (string.IsNullOrWhiteSpace(initialPath))
            {
                var fallback = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                if (Directory.Exists(fallback))
                {
                    initialPath = fallback;
                }
            }

            if (!string.IsNullOrWhiteSpace(initialPath))
            {
                await CreateAndOpenTabAsync(initialPath);
            }
            else
            {
                AddLog("利用可能なフォルダーが見つかりません。");
            }
        }
        catch (Exception exception)
        {
            ShowError("起動処理に失敗しました。", exception);
        }
    }

    private void LoadDrives()
    {
        _drives.Clear();
        foreach (var drive in _libraryService.GetAvailableDrives())
        {
            _drives.Add(drive);
        }
    }

    private void LoadDirectoryRoots()
    {
        _directoryRoots.Clear();
        foreach (var drive in _drives)
        {
            _directoryRoots.Add(CreateDirectoryNode(drive.Path, drive.Label));
        }
    }

    private DirectoryTreeNode CreateDirectoryNode(string path, string label)
    {
        var node = new DirectoryTreeNode
        {
            Path = path,
            Label = label
        };

        if (HasChildDirectories(path))
        {
            node.Children.Add(DirectoryTreeNode.CreatePlaceholder());
        }

        return node;
    }

    private void LoadChildNodes(DirectoryTreeNode node)
    {
        if (node.AreChildrenLoaded || node.IsPlaceholder)
        {
            return;
        }

        node.Children.Clear();
        foreach (var childPath in EnumerateDirectoriesSafe(node.Path))
        {
            var name = Path.GetFileName(childPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            node.Children.Add(CreateDirectoryNode(childPath, string.IsNullOrWhiteSpace(name) ? childPath : name));
        }

        node.AreChildrenLoaded = true;
    }

    private static IEnumerable<string> EnumerateDirectoriesSafe(string path)
    {
        try
        {
            var options = new EnumerationOptions
            {
                AttributesToSkip = 0,
                IgnoreInaccessible = true,
                RecurseSubdirectories = false,
                ReturnSpecialDirectories = false
            };

            return Directory.EnumerateDirectories(path, "*", options)
                .OrderBy(child => Path.GetFileName(child), StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
        catch
        {
            return [];
        }
    }

    private static bool HasChildDirectories(string path)
    {
        try
        {
            var options = new EnumerationOptions
            {
                AttributesToSkip = 0,
                IgnoreInaccessible = true,
                RecurseSubdirectories = false,
                ReturnSpecialDirectories = false
            };

            return Directory.EnumerateDirectories(path, "*", options).Any();
        }
        catch
        {
            return false;
        }
    }

    private void SelectDirectoryNodeForPath(string? path)
    {
        if (_suppressTreeSelection || string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var rootPath = Path.GetPathRoot(path);
        if (string.IsNullOrWhiteSpace(rootPath))
        {
            return;
        }

        var rootNode = _directoryRoots.FirstOrDefault(node =>
            string.Equals(Path.GetPathRoot(node.Path) ?? node.Path, rootPath, StringComparison.OrdinalIgnoreCase));

        if (rootNode is null)
        {
            return;
        }

        _suppressTreeSelection = true;
        if (_selectedDirectoryNode is not null)
        {
            _selectedDirectoryNode.IsSelected = false;
        }

        var current = rootNode;
        LoadChildNodes(current);

        var relative = Path.GetRelativePath(rootPath, path);
        if (!string.Equals(relative, ".", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var segment in relative.Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries))
            {
                current.IsExpanded = true;
                LoadChildNodes(current);
                var nextPath = Path.Combine(current.Path, segment);
                var child = current.Children.FirstOrDefault(node => string.Equals(node.Path, nextPath, StringComparison.OrdinalIgnoreCase));
                if (child is null)
                {
                    break;
                }

                current = child;
            }
        }

        current.IsExpanded = true;
        current.IsSelected = true;
        _selectedDirectoryNode = current;
        _suppressTreeSelection = false;
    }

    private async Task<ExplorerTabState> CreateAndOpenTabAsync(string path)
    {
        return await CreateAndOpenTabAsync(path, _selectedWorkspace);
    }

    private async Task<ExplorerTabState> CreateAndOpenTabAsync(string path, WorkspacePaneState? workspace)
    {
        var tab = new ExplorerTabState();
        tab.AssetsView.Filter = item => FilterAsset(tab, item);
        tab.PropertyChanged += Tab_PropertyChanged;
        ApplyGrouping(tab);
        _tabs.Add(tab);
        var targetWorkspace = workspace ?? GetFallbackWorkspace();
        AttachTabToWorkspace(targetWorkspace, tab);
        _selectedWorkspace = targetWorkspace;
        SelectedTab = tab;
        await NavigateToPathAsync(tab, path);
        return tab;
    }

    private async void NewTab_Click(object sender, RoutedEventArgs e)
    {
        var path = SelectedTab?.CurrentLocationPath ?? _drives.FirstOrDefault()?.Path;
        if (!string.IsNullOrWhiteSpace(path))
        {
            await CreateAndOpenTabAsync(path);
        }
    }

    private async void NewTabFromContext_Click(object sender, RoutedEventArgs e)
    {
        var path = SelectedTab?.CurrentLocationPath ?? _drives.FirstOrDefault()?.Path;
        if (!string.IsNullOrWhiteSpace(path))
        {
            await CreateAndOpenTabAsync(path);
        }
    }

    private async void CloseTab_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.Tag is ExplorerTabState tab)
        {
            await CloseTabAsync(tab);
        }
    }

    private async Task CloseTabAsync(ExplorerTabState tab)
    {
        tab.PropertyChanged -= Tab_PropertyChanged;
        tab.LoadCts?.Cancel();

        var sourceWorkspace = FindWorkspaceForTab(tab);
        var removedIndex = _tabs.IndexOf(tab);
        _tabs.Remove(tab);
        DetachTabFromWorkspace(sourceWorkspace, tab);
        NormalizeWorkspaceLayout();

        if (_tabs.Count == 0)
        {
            var fallbackPath = _drives.FirstOrDefault()?.Path;
            if (!string.IsNullOrWhiteSpace(fallbackPath))
            {
                await CreateAndOpenTabAsync(fallbackPath);
            }
            else
            {
                SelectedTab = null;
            }

            return;
        }

        var nextTab = sourceWorkspace?.SelectedTab ?? _tabs[Math.Clamp(removedIndex - 1, 0, _tabs.Count - 1)];
        SelectedTab = nextTab;
    }

    private void DirectoryTreeItem_Expanded(object sender, RoutedEventArgs e)
    {
        if (e.OriginalSource is TreeViewItem { DataContext: DirectoryTreeNode node })
        {
            LoadChildNodes(node);
        }
    }

    private async void DirectoryTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (_suppressTreeSelection || e.NewValue is not DirectoryTreeNode node || string.IsNullOrWhiteSpace(node.Path))
        {
            return;
        }

        _selectedDirectoryNode = node;
        if (SelectedTab is not null)
        {
            await NavigateToPathAsync(SelectedTab, node.Path);
        }
    }
    private async void Back_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTab is not { CanGoBack: true } tab)
        {
            return;
        }

        tab.HistoryIndex -= 1;
        await NavigateToPathAsync(tab, tab.History[tab.HistoryIndex], addToHistory: false);
    }

    private async void Forward_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTab is not { CanGoForward: true } tab)
        {
            return;
        }

        tab.HistoryIndex += 1;
        await NavigateToPathAsync(tab, tab.History[tab.HistoryIndex], addToHistory: false);
    }

    private async void Up_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTab is not { CanGoUp: true } tab)
        {
            return;
        }

        var parent = Directory.GetParent(tab.CurrentLocationPath);
        if (parent is not null)
        {
            await NavigateToPathAsync(tab, parent.FullName);
        }
    }

    private async void RenameSelected_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTab?.SelectedAsset is AssetItem asset)
        {
            await RenameAssetAsync(SelectedTab, asset);
        }
    }

    private void Toolbar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left || IsInteractiveElement(e.OriginalSource as DependencyObject))
        {
            return;
        }

        if (e.ClickCount == 2)
        {
            ToggleWindowState();
            return;
        }

        DragMove();
    }

    private void Minimize_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    private void ToggleMaximize_Click(object sender, RoutedEventArgs e)
    {
        ToggleWindowState();
    }

    private void CloseWindow_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ToggleWindowState()
    {
        if (ResizeMode == ResizeMode.NoResize)
        {
            return;
        }

        WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
    }

    private async void AssetListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: ExplorerTabState tab } && tab.SelectedAsset is not null)
        {
            await OpenAssetAsync(tab, tab.SelectedAsset);
        }
    }

    private async void AssetGrid_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Middle ||
            e.OriginalSource is not DependencyObject source ||
            sender is not DataGrid { DataContext: ExplorerTabState tab })
        {
            return;
        }

        var row = FindAncestor<DataGridRow>(source);
        if (row?.Item is not AssetItem asset || !asset.IsDirectory)
        {
            return;
        }

        row.IsSelected = true;
        row.Focus();
        tab.SelectedAsset = asset;

        var sourceWorkspace = FindWorkspaceForTab(tab) ?? GetFallbackWorkspace();
        var destination = GetAdjacentWorkspace(sourceWorkspace, WorkspaceDropZone.Right) ?? sourceWorkspace;
        await CreateAndOpenTabAsync(asset.FilePath, destination);
        e.Handled = true;
    }

    private void AssetGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.OriginalSource is not DependencyObject source)
        {
            return;
        }

        var row = FindAncestor<DataGridRow>(source);
        if (row is not null)
        {
            row.IsSelected = true;
            row.Focus();
        }
    }

    private void AssetGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (sender is not DataGrid dataGrid || dataGrid.DataContext is not ExplorerTabState tab)
        {
            return;
        }

        var source = e.OriginalSource as DependencyObject ?? Mouse.DirectlyOver as DependencyObject;
        var row = FindAncestor<DataGridRow>(source);
        if (row?.Item is AssetItem asset)
        {
            tab.SelectedAsset = asset;
            dataGrid.ContextMenu = CreateAssetContextMenu(tab, asset);
            return;
        }

        dataGrid.ContextMenu = CreateGridContextMenu();
    }

    private async void OpenContextAsset_Click(object sender, RoutedEventArgs e)
    {
        if (TryGetContextAsset(sender, out var tab, out var asset))
        {
            await OpenAssetAsync(tab, asset);
        }
    }

    private async void OpenInNewTab_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetContextAsset(sender, out var ownerTab, out var asset))
        {
            return;
        }

        var path = asset.IsDirectory ? asset.FilePath : Path.GetDirectoryName(asset.FilePath);
        if (!string.IsNullOrWhiteSpace(path))
        {
            await CreateAndOpenTabAsync(path, FindWorkspaceForTab(ownerTab));
        }
    }

    private void CopySelected_Click(object sender, RoutedEventArgs e)
    {
        if (TryGetContextAsset(sender, out var tab, out var asset))
        {
            BeginTransfer(tab, asset, isMove: false);
        }
    }

    private void CutSelected_Click(object sender, RoutedEventArgs e)
    {
        if (TryGetContextAsset(sender, out var tab, out var asset))
        {
            BeginTransfer(tab, asset, isMove: true);
        }
    }

    private async void PasteIntoCurrentDirectory_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTab is not null)
        {
            await PasteTransferAsync(SelectedTab.CurrentLocationPath);
        }
    }

    private async void PasteIntoContextFolder_Click(object sender, RoutedEventArgs e)
    {
        if (TryGetContextAsset(sender, out _, out var asset) && asset.IsDirectory)
        {
            await PasteTransferAsync(asset.FilePath);
        }
    }

    private void CopyPath_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetContextAsset(sender, out _, out var asset))
        {
            return;
        }

        Clipboard.SetText(asset.FilePath);
        AddLog($"{asset.Name} のパスをコピーしました。");
    }

    private void ShowInExplorer_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetContextAsset(sender, out _, out var asset))
        {
            return;
        }

        ShowInExplorer(asset.FilePath);
    }

    private async void RenameContextAsset_Click(object sender, RoutedEventArgs e)
    {
        if (TryGetContextAsset(sender, out var tab, out var asset))
        {
            await RenameAssetAsync(tab, asset);
        }
    }

    private async void DeleteContextAsset_Click(object sender, RoutedEventArgs e)
    {
        if (TryGetContextAsset(sender, out var tab, out var asset))
        {
            await DeleteAssetAsync(tab, asset);
        }
    }

    private async void DeleteSelected_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTab?.SelectedAsset is AssetItem asset)
        {
            await DeleteAssetAsync(SelectedTab, asset);
        }
    }

    private async void AddTag_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: ExplorerTabState tab } || tab.SelectedAsset is null)
        {
            return;
        }

        var tags = tab.PendingTagText
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (tags.Length == 0)
        {
            return;
        }

        tab.SelectedAsset.Tags = tab.SelectedAsset.Tags
            .Concat(tags)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(tag => tag, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        tab.PendingTagText = string.Empty;
        UpsertMetadata(tab.MetadataDocument, tab.SelectedAsset);
        tab.RefreshDerivedState();
        await PersistMetadataAsync(tab);
        AddLog($"{tab.SelectedAsset.Name} にタグを追加しました。");
    }

    private async void RemoveTag_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: string tag } element || FindTabFromVisual(element) is not ExplorerTabState tab || tab.SelectedAsset is null)
        {
            return;
        }

        tab.SelectedAsset.Tags = tab.SelectedAsset.Tags
            .Where(existing => !string.Equals(existing, tag, StringComparison.OrdinalIgnoreCase))
            .ToArray();

        UpsertMetadata(tab.MetadataDocument, tab.SelectedAsset);
        tab.RefreshDerivedState();
        await PersistMetadataAsync(tab);
        AddLog($"{tab.SelectedAsset.Name} からタグ {tag} を削除しました。");
    }

    private async void FavoriteToggle_Changed(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: ExplorerTabState tab } || tab.SelectedAsset is null)
        {
            return;
        }

        UpsertMetadata(tab.MetadataDocument, tab.SelectedAsset);
        tab.RefreshDerivedState();
        await PersistMetadataAsync(tab);
        AddLog($"{tab.SelectedAsset.Name} のお気に入り設定を更新しました。");
    }

    private async void NoteEditor_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: ExplorerTabState tab } || tab.SelectedAsset is null)
        {
            return;
        }

        UpsertMetadata(tab.MetadataDocument, tab.SelectedAsset);
        await PersistMetadataAsync(tab);
        AddLog($"{tab.SelectedAsset.Name} のメモを保存しました。");
    }

    private void Tab_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is not ExplorerTabState tab)
        {
            return;
        }

        if (e.PropertyName == nameof(ExplorerTabState.SelectedAsset))
        {
            _ = LoadPreviewAsync(tab, tab.SelectedAsset);
            return;
        }

        if (e.PropertyName is nameof(ExplorerTabState.SearchText) or nameof(ExplorerTabState.SelectedTypeFilter) or nameof(ExplorerTabState.FavoritesOnly))
        {
            tab.AssetsView.Refresh();
            tab.NotifyProperty(nameof(ExplorerTabState.ResultSummary));
            return;
        }

        if (e.PropertyName == nameof(ExplorerTabState.SelectedGroupingOption))
        {
            ApplyGrouping(tab);
        }
    }
    private async Task OpenAssetAsync(ExplorerTabState tab, AssetItem asset)
    {
        if (asset.IsDirectory)
        {
            await NavigateToPathAsync(tab, asset.FilePath);
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = asset.FilePath,
            UseShellExecute = true
        });
    }

    private async Task NavigateToPathAsync(ExplorerTabState tab, string directoryPath, string? preferredSelectionPath = null, bool addToHistory = true)
    {
        if (!Directory.Exists(directoryPath))
        {
            return;
        }

        tab.LoadCts?.Cancel();
        tab.LoadCts = new CancellationTokenSource();
        var cancellationToken = tab.LoadCts.Token;

        try
        {
            tab.StatusText = "フォルダーを読み込み中...";
            tab.CurrentLocationPath = directoryPath;
            tab.CurrentDrivePath = Path.GetPathRoot(directoryPath) ?? directoryPath;
            tab.UpdateTitle();

            tab.MetadataDocument = await GetMetadataDocumentForRootAsync(tab.CurrentDrivePath, cancellationToken);
            var entries = await _libraryService.GetDirectoryEntriesAsync(tab.CurrentDrivePath, directoryPath, tab.MetadataDocument.Items, cancellationToken);

            tab.Assets.Clear();
            foreach (var entry in entries)
            {
                tab.Assets.Add(entry);
            }

            tab.SuppressNavigationSelection = true;
            tab.SelectedDrive = _drives.FirstOrDefault(drive => string.Equals(drive.Path, tab.CurrentDrivePath, StringComparison.OrdinalIgnoreCase));
            tab.ChildFolders.Clear();
            foreach (var folder in _libraryService.GetChildFolders(directoryPath))
            {
                tab.ChildFolders.Add(folder);
            }

            tab.SelectedChildFolder = null;
            tab.SuppressNavigationSelection = false;

            if (addToHistory)
            {
                PushHistory(tab, directoryPath);
            }
            else
            {
                tab.NotifyProperty(nameof(ExplorerTabState.CanGoBack));
                tab.NotifyProperty(nameof(ExplorerTabState.CanGoForward));
            }

            ApplyGrouping(tab);
            tab.RefreshDerivedState();
            tab.SelectedAsset = preferredSelectionPath is null
                ? null
                : tab.Assets.FirstOrDefault(asset => string.Equals(asset.FilePath, preferredSelectionPath, StringComparison.OrdinalIgnoreCase));

            SelectDirectoryNodeForPath(directoryPath);
            tab.StatusText = $"{tab.Assets.Count} 件を表示中";
            AddLog($"{directoryPath} を開きました。");
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception exception)
        {
            ShowError("フォルダーの読み込みに失敗しました。", exception);
        }
    }

    private async Task LoadPreviewAsync(ExplorerTabState tab, AssetItem? asset)
    {
        tab.ResetPreview();

        if (asset is null)
        {
            return;
        }

        try
        {
            switch (asset.Kind)
            {
                case AssetKind.Folder:
                    tab.PreviewFallbackText = "フォルダーです。ダブルクリックで中を開けます。";
                    tab.IsFallbackPreviewVisible = true;
                    break;
                case AssetKind.Image:
                    tab.PreviewImagePath = asset.FilePath;
                    tab.IsImagePreviewVisible = true;
                    break;
                case AssetKind.Video:
                    tab.PreviewVideoUri = new Uri(asset.FilePath);
                    tab.IsVideoPreviewVisible = true;
                    break;
                case AssetKind.Text:
                    tab.PreviewText = await File.ReadAllTextAsync(asset.FilePath);
                    if (tab.PreviewText.Length > 20000)
                    {
                        tab.PreviewText = tab.PreviewText[..20000];
                    }

                    tab.IsTextPreviewVisible = true;
                    break;
                case AssetKind.Pdf:
                    tab.PreviewFallbackText = "PDF は外部ビューアーで開いて確認してください。";
                    tab.IsFallbackPreviewVisible = true;
                    break;
                default:
                    tab.PreviewFallbackText = "この形式はプレビューに対応していません。";
                    tab.IsFallbackPreviewVisible = true;
                    break;
            }
        }
        catch (Exception exception)
        {
            tab.PreviewFallbackText = $"プレビューの読み込みに失敗しました: {exception.Message}";
            tab.IsFallbackPreviewVisible = true;
        }
    }

    private bool FilterAsset(ExplorerTabState tab, object item)
    {
        if (item is not AssetItem asset)
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(tab.SearchText))
        {
            var search = tab.SearchText.Trim();
            var matchesSearch = asset.Name.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                                asset.Tags.Any(tag => tag.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
                                asset.Note.Contains(search, StringComparison.OrdinalIgnoreCase);
            if (!matchesSearch)
            {
                return false;
            }
        }

        if (tab.FavoritesOnly && !asset.IsFavorite)
        {
            return false;
        }

        return tab.SelectedTypeFilter switch
        {
            "フォルダー" => asset.Kind == AssetKind.Folder,
            "画像" => asset.Kind == AssetKind.Image,
            "動画" => asset.Kind == AssetKind.Video,
            "テキスト" => asset.Kind == AssetKind.Text,
            "PDF" => asset.Kind == AssetKind.Pdf,
            "その他" => asset.Kind == AssetKind.Other,
            _ => true
        };
    }

    private async Task RenameAssetAsync(ExplorerTabState tab, AssetItem asset)
    {
        var nextName = Interaction.InputBox("新しい名前を入力してください。", "名前変更", asset.Name);
        if (string.IsNullOrWhiteSpace(nextName) || string.Equals(nextName, asset.Name, StringComparison.Ordinal))
        {
            return;
        }

        var targetPath = Path.Combine(Path.GetDirectoryName(asset.FilePath)!, nextName);
        if (File.Exists(targetPath) || Directory.Exists(targetPath))
        {
            MessageBox.Show("同じ名前の項目が既に存在します。", "名前変更", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            if (asset.IsDirectory)
            {
                Directory.Move(asset.FilePath, targetPath);
            }
            else
            {
                File.Move(asset.FilePath, targetPath);
            }

            await TransferMetadataAsync(asset.FilePath, targetPath, removeSource: true);
            AddLog($"{asset.Name} を {nextName} に変更しました。");
            await RefreshAffectedTabsAsync([Path.GetDirectoryName(asset.FilePath)!, Path.GetDirectoryName(targetPath)!], targetPath);
        }
        catch (Exception exception)
        {
            ShowError("名前変更に失敗しました。", exception);
        }
    }

    private async Task DeleteAssetAsync(ExplorerTabState tab, AssetItem asset)
    {
        var confirmed = MessageBox.Show($"{asset.Name} をごみ箱へ移動しますか？", "削除", MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (confirmed != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            if (asset.IsDirectory)
            {
                Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(asset.FilePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
            }
            else
            {
                Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(asset.FilePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
            }

            RemoveMetadataPrefix(tab.MetadataDocument, asset.FilePath);
            await PersistMetadataAsync(tab);
            AddLog($"{asset.Name} を削除しました。");
            await RefreshAffectedTabsAsync([tab.CurrentLocationPath], null);
        }
        catch (Exception exception)
        {
            ShowError("削除に失敗しました。", exception);
        }
    }

    private void BeginTransfer(ExplorerTabState tab, AssetItem asset, bool isMove)
    {
        _pendingTransfer = new PendingTransferOperation
        {
            SourceRootPath = tab.CurrentDrivePath,
            SourcePaths = [asset.FilePath],
            IsMove = isMove
        };

        AddLog($"{asset.Name} を{(isMove ? "切り取り" : "コピー")}待ちにしました。");
        OnPropertyChanged(nameof(HasPendingTransfer));
    }

    private async Task PasteTransferAsync(string destinationDirectory)
    {
        if (_pendingTransfer is null || !Directory.Exists(destinationDirectory))
        {
            return;
        }

        try
        {
            var destinationPaths = new List<(string SourcePath, string DestinationPath)>();

            foreach (var sourcePath in _pendingTransfer.SourcePaths)
            {
                var itemName = Path.GetFileName(sourcePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                var destinationPath = Path.Combine(destinationDirectory, itemName);

                if (string.Equals(sourcePath, destinationPath, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("同じ場所には貼り付けできません。", "貼り付け", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (IsChildPath(sourcePath, destinationDirectory))
                {
                    MessageBox.Show("フォルダーをその子フォルダーへ移動またはコピーすることはできません。", "貼り付け", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (File.Exists(destinationPath) || Directory.Exists(destinationPath))
                {
                    MessageBox.Show($"同名の項目が既に存在します: {itemName}", "貼り付け", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                destinationPaths.Add((sourcePath, destinationPath));
            }

            foreach (var (sourcePath, destinationPath) in destinationPaths)
            {
                TransferPath(sourcePath, destinationPath, _pendingTransfer.IsMove);
                await TransferMetadataAsync(sourcePath, destinationPath, removeSource: _pendingTransfer.IsMove);
                AddLog($"{Path.GetFileName(sourcePath)} を{(_pendingTransfer.IsMove ? "移動" : "コピー")}しました。");
            }

            if (_pendingTransfer.IsMove)
            {
                _pendingTransfer = null;
                OnPropertyChanged(nameof(HasPendingTransfer));
            }

            var refreshTargets = destinationPaths
                .SelectMany(pair => new[] { Path.GetDirectoryName(pair.SourcePath)!, destinationDirectory })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            await RefreshAffectedTabsAsync(refreshTargets, null);
        }
        catch (Exception exception)
        {
            ShowError("貼り付けに失敗しました。", exception);
        }
    }
    private static void TransferPath(string sourcePath, string destinationPath, bool isMove)
    {
        if (Directory.Exists(sourcePath))
        {
            if (isMove)
            {
                MoveDirectory(sourcePath, destinationPath);
            }
            else
            {
                CopyDirectory(sourcePath, destinationPath);
            }

            return;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
        if (isMove)
        {
            MoveFile(sourcePath, destinationPath);
        }
        else
        {
            File.Copy(sourcePath, destinationPath);
        }
    }

    private async Task RefreshAffectedTabsAsync(IEnumerable<string> paths, string? preferredSelectionPath)
    {
        var affectedPaths = paths.Where(path => !string.IsNullOrWhiteSpace(path)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

        foreach (var tab in _tabs.ToArray())
        {
            if (!Directory.Exists(tab.CurrentLocationPath))
            {
                var fallback = Directory.GetParent(tab.CurrentLocationPath)?.FullName ?? _drives.FirstOrDefault()?.Path;
                if (!string.IsNullOrWhiteSpace(fallback) && Directory.Exists(fallback))
                {
                    await NavigateToPathAsync(tab, fallback, addToHistory: false);
                }

                continue;
            }

            if (affectedPaths.Contains(tab.CurrentLocationPath, StringComparer.OrdinalIgnoreCase))
            {
                await NavigateToPathAsync(tab, tab.CurrentLocationPath, preferredSelectionPath, addToHistory: false);
            }
        }
    }

    private async Task<LibraryMetadataDocument> GetMetadataDocumentForRootAsync(string rootPath, CancellationToken cancellationToken)
    {
        var existingTab = _tabs.FirstOrDefault(tab =>
            !ReferenceEquals(tab, SelectedTab) && string.Equals(tab.CurrentDrivePath, rootPath, StringComparison.OrdinalIgnoreCase));

        return existingTab is not null ? existingTab.MetadataDocument : await _metadataStore.LoadAsync(rootPath, cancellationToken);
    }

    private async Task PersistMetadataAsync(ExplorerTabState tab)
    {
        if (!string.IsNullOrWhiteSpace(tab.CurrentDrivePath))
        {
            await _metadataStore.SaveAsync(tab.CurrentDrivePath, tab.MetadataDocument, CancellationToken.None);
        }
    }

    private async Task TransferMetadataAsync(string sourcePath, string destinationPath, bool removeSource)
    {
        var sourceRoot = Path.GetPathRoot(sourcePath) ?? sourcePath;
        var destinationRoot = Path.GetPathRoot(destinationPath) ?? destinationPath;
        var sourceDocument = await GetOpenOrStoredMetadataAsync(sourceRoot);
        var destinationDocument = string.Equals(sourceRoot, destinationRoot, StringComparison.OrdinalIgnoreCase)
            ? sourceDocument
            : await GetOpenOrStoredMetadataAsync(destinationRoot);

        foreach (var (key, value) in CloneMetadataPrefix(sourceDocument, sourcePath, destinationPath))
        {
            destinationDocument.Items[key] = value;
        }

        if (removeSource)
        {
            RemoveMetadataPrefix(sourceDocument, sourcePath);
        }

        await _metadataStore.SaveAsync(sourceRoot, sourceDocument, CancellationToken.None);
        if (!string.Equals(sourceRoot, destinationRoot, StringComparison.OrdinalIgnoreCase))
        {
            await _metadataStore.SaveAsync(destinationRoot, destinationDocument, CancellationToken.None);
        }

        foreach (var tab in _tabs.Where(tab => string.Equals(tab.CurrentDrivePath, sourceRoot, StringComparison.OrdinalIgnoreCase)))
        {
            tab.MetadataDocument = sourceDocument;
        }

        foreach (var tab in _tabs.Where(tab => string.Equals(tab.CurrentDrivePath, destinationRoot, StringComparison.OrdinalIgnoreCase)))
        {
            tab.MetadataDocument = destinationDocument;
        }
    }

    private async Task<LibraryMetadataDocument> GetOpenOrStoredMetadataAsync(string rootPath)
    {
        var existing = _tabs.FirstOrDefault(tab => string.Equals(tab.CurrentDrivePath, rootPath, StringComparison.OrdinalIgnoreCase));
        return existing is not null ? existing.MetadataDocument : await _metadataStore.LoadAsync(rootPath, CancellationToken.None);
    }

    private static Dictionary<string, AssetMetadata> CloneMetadataPrefix(LibraryMetadataDocument sourceDocument, string sourcePath, string destinationPath)
    {
        return sourceDocument.Items
            .Where(pair => pair.Key.Equals(sourcePath, StringComparison.OrdinalIgnoreCase) || pair.Key.StartsWith(sourcePath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(
                pair => pair.Key.Equals(sourcePath, StringComparison.OrdinalIgnoreCase) ? destinationPath : destinationPath + pair.Key[sourcePath.Length..],
                pair => new AssetMetadata { IsFavorite = pair.Value.IsFavorite, Note = pair.Value.Note, Tags = [.. pair.Value.Tags] },
                StringComparer.OrdinalIgnoreCase);
    }

    private static void UpsertMetadata(LibraryMetadataDocument document, AssetItem asset)
    {
        document.Items[asset.FilePath] = new AssetMetadata
        {
            Tags = [.. asset.Tags],
            Note = asset.Note,
            IsFavorite = asset.IsFavorite
        };
    }

    private static void RemoveMetadataPrefix(LibraryMetadataDocument document, string path)
    {
        var keys = document.Items.Keys
            .Where(existing => existing.Equals(path, StringComparison.OrdinalIgnoreCase) || existing.StartsWith(path + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .ToArray();

        foreach (var key in keys)
        {
            document.Items.Remove(key);
        }
    }

    private static void PushHistory(ExplorerTabState tab, string directoryPath)
    {
        if (tab.HistoryIndex >= 0 && string.Equals(tab.History[tab.HistoryIndex], directoryPath, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        while (tab.History.Count - 1 > tab.HistoryIndex)
        {
            tab.History.RemoveAt(tab.History.Count - 1);
        }

        tab.History.Add(directoryPath);
        tab.HistoryIndex = tab.History.Count - 1;
    }

    private static bool IsChildPath(string sourcePath, string destinationDirectory)
    {
        if (!Directory.Exists(sourcePath))
        {
            return false;
        }

        var normalizedSource = Path.GetFullPath(sourcePath).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        var normalizedDestination = Path.GetFullPath(destinationDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        return normalizedDestination.StartsWith(normalizedSource, StringComparison.OrdinalIgnoreCase);
    }

    private static void CopyDirectory(string sourceDirectory, string destinationDirectory)
    {
        Directory.CreateDirectory(destinationDirectory);
        foreach (var file in Directory.EnumerateFiles(sourceDirectory))
        {
            File.Copy(file, Path.Combine(destinationDirectory, Path.GetFileName(file)));
        }

        foreach (var directory in Directory.EnumerateDirectories(sourceDirectory))
        {
            CopyDirectory(directory, Path.Combine(destinationDirectory, Path.GetFileName(directory)));
        }
    }

    private static void MoveDirectory(string sourceDirectory, string destinationDirectory)
    {
        var sourceRoot = Path.GetPathRoot(sourceDirectory);
        var destinationRoot = Path.GetPathRoot(destinationDirectory);
        if (string.Equals(sourceRoot, destinationRoot, StringComparison.OrdinalIgnoreCase))
        {
            Directory.Move(sourceDirectory, destinationDirectory);
            return;
        }

        CopyDirectory(sourceDirectory, destinationDirectory);
        Directory.Delete(sourceDirectory, recursive: true);
    }

    private static void MoveFile(string sourceFile, string destinationFile)
    {
        var sourceRoot = Path.GetPathRoot(sourceFile);
        var destinationRoot = Path.GetPathRoot(destinationFile);
        if (string.Equals(sourceRoot, destinationRoot, StringComparison.OrdinalIgnoreCase))
        {
            File.Move(sourceFile, destinationFile);
            return;
        }

        File.Copy(sourceFile, destinationFile);
        File.Delete(sourceFile);
    }

    private void ShowInExplorer(string filePath)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"/select,\"{filePath}\"",
            UseShellExecute = true
        });

        AddLog($"{Path.GetFileName(filePath)} を Explorer で表示しました。");
    }

    private bool TryGetContextAsset(object sender, out ExplorerTabState tab, out AssetItem asset)
    {
        tab = null!;
        asset = null!;

        if (sender is not MenuItem { Tag: AssetItem tagAsset, CommandParameter: ExplorerTabState ownerTab })
        {
            return false;
        }

        tab = ownerTab;
        asset = tagAsset;
        tab.SelectedAsset = asset;
        return true;
    }

    private ContextMenu CreateGridContextMenu()
    {
        var menu = new ContextMenu();
        var newTabItem = new MenuItem { Header = "新しいタブ" };
        newTabItem.Click += NewTabFromContext_Click;
        menu.Items.Add(newTabItem);

        var pasteItem = new MenuItem { Header = "ここに貼り付け", IsEnabled = HasPendingTransfer };
        pasteItem.Click += PasteIntoCurrentDirectory_Click;
        menu.Items.Add(pasteItem);
        return menu;
    }

    private ContextMenu CreateAssetContextMenu(ExplorerTabState tab, AssetItem asset)
    {
        var menu = new ContextMenu();
        menu.Items.Add(CreateAssetMenuItem("開く", asset, tab, OpenContextAsset_Click));
        menu.Items.Add(CreateAssetMenuItem("新しいタブで開く", asset, tab, OpenInNewTab_Click));
        menu.Items.Add(new Separator());
        menu.Items.Add(CreateAssetMenuItem("コピー", asset, tab, CopySelected_Click));
        menu.Items.Add(CreateAssetMenuItem("切り取り", asset, tab, CutSelected_Click));

        var pasteIntoItem = CreateAssetMenuItem("このフォルダーに貼り付け", asset, tab, PasteIntoContextFolder_Click);
        pasteIntoItem.IsEnabled = asset.IsDirectory && HasPendingTransfer;
        menu.Items.Add(pasteIntoItem);

        menu.Items.Add(new Separator());
        menu.Items.Add(CreateAssetMenuItem("名前変更", asset, tab, RenameContextAsset_Click));
        menu.Items.Add(CreateAssetMenuItem("Explorer で表示", asset, tab, ShowInExplorer_Click));
        menu.Items.Add(CreateAssetMenuItem("パスをコピー", asset, tab, CopyPath_Click));
        menu.Items.Add(CreateAssetMenuItem("削除", asset, tab, DeleteContextAsset_Click));
        return menu;
    }

    private static MenuItem CreateAssetMenuItem(string header, AssetItem asset, ExplorerTabState tab, RoutedEventHandler handler)
    {
        var item = new MenuItem { Header = header, Tag = asset, CommandParameter = tab };
        item.Click += handler;
        return item;
    }

    private void WorkspacePane_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is not WorkspacePaneState pane)
        {
            return;
        }

        if (e.PropertyName == nameof(WorkspacePaneState.SelectedTab) && pane.SelectedTab is not null)
        {
            _selectedWorkspace = pane;
            if (!ReferenceEquals(_selectedTab, pane.SelectedTab))
            {
                _selectedTab = pane.SelectedTab;
                OnPropertyChanged(nameof(SelectedTab));
                SelectDirectoryNodeForPath(_selectedTab.CurrentLocationPath);
            }
        }

        if (e.PropertyName is nameof(WorkspacePaneState.SelectedTab) or nameof(WorkspacePaneState.IsVisible) or nameof(WorkspacePaneState.HasTabs))
        {
            UpdateWorkspaceLayout();
        }
    }

    private WorkspacePaneState GetFallbackWorkspace()
    {
        return _selectedWorkspace
               ?? _workspacePanes.FirstOrDefault(pane => pane.IsVisible)
               ?? _workspacePanes[0];
    }

    private WorkspacePaneState? FindWorkspaceForTab(ExplorerTabState tab)
    {
        return _workspacePanes.FirstOrDefault(pane => pane.Tabs.Contains(tab));
    }

    private void AttachTabToWorkspace(WorkspacePaneState workspace, ExplorerTabState tab)
    {
        workspace.IsVisible = true;
        if (!workspace.Tabs.Contains(tab))
        {
            workspace.Tabs.Add(tab);
        }

        workspace.SelectedTab = tab;
        UpdateWorkspaceLayout();
    }

    private void DetachTabFromWorkspace(WorkspacePaneState? workspace, ExplorerTabState tab)
    {
        if (workspace is null)
        {
            return;
        }

        workspace.Tabs.Remove(tab);
        if (!workspace.HasTabs && workspace.SlotIndex != 0)
        {
            workspace.IsVisible = false;
        }

        UpdateWorkspaceLayout();
    }

    private void MoveAllTabs(WorkspacePaneState source, WorkspacePaneState destination)
    {
        var tabs = source.Tabs.ToArray();
        foreach (var tab in tabs)
        {
            source.Tabs.Remove(tab);
            if (!destination.Tabs.Contains(tab))
            {
                destination.Tabs.Add(tab);
            }
        }

        destination.IsVisible = true;
        destination.SelectedTab = source.SelectedTab ?? destination.SelectedTab ?? destination.Tabs.FirstOrDefault();
        source.SelectedTab = null;
        if (source.SlotIndex != 0)
        {
            source.IsVisible = false;
        }
    }

    private void NormalizeWorkspaceLayout()
    {
        if (!Workspace0.HasTabs)
        {
            var nextWorkspace = _workspacePanes.Skip(1).FirstOrDefault(pane => pane.HasTabs);
            if (nextWorkspace is not null)
            {
                MoveAllTabs(nextWorkspace, Workspace0);
            }
        }

        if (!Workspace1.HasTabs && Workspace3.HasTabs && !Workspace2.HasTabs)
        {
            MoveAllTabs(Workspace3, Workspace1);
        }

        if (!Workspace2.HasTabs && Workspace3.HasTabs && !Workspace1.HasTabs)
        {
            MoveAllTabs(Workspace3, Workspace2);
        }

        foreach (var pane in _workspacePanes.Where(pane => pane.SlotIndex != 0 && !pane.HasTabs))
        {
            pane.IsVisible = false;
            pane.SelectedTab = null;
        }

        Workspace0.IsVisible = true;
        _selectedWorkspace = _selectedTab is not null ? FindWorkspaceForTab(_selectedTab) ?? Workspace0 : GetFallbackWorkspace();
        UpdateWorkspaceLayout();
    }

    private void MoveTabToWorkspace(ExplorerTabState tab, WorkspacePaneState destination)
    {
        var source = FindWorkspaceForTab(tab);
        if (source is not null && ReferenceEquals(source, destination))
        {
            destination.SelectedTab = tab;
            SelectedTab = tab;
            return;
        }

        if (source is not null)
        {
            source.Tabs.Remove(tab);
            if (!source.HasTabs && source.SlotIndex != 0)
            {
                source.IsVisible = false;
            }
        }

        destination.IsVisible = true;
        if (!destination.Tabs.Contains(tab))
        {
            destination.Tabs.Add(tab);
        }

        destination.SelectedTab = tab;
        _selectedWorkspace = destination;
        SelectedTab = tab;
        NormalizeWorkspaceLayout();
    }

    private void UpdateWorkspaceLayout()
    {
        if (!IsInitialized)
        {
            return;
        }

        var hasRightColumn = Workspace1.HasTabs || Workspace3.HasTabs;
        var hasBottomRow = Workspace2.HasTabs || Workspace3.HasTabs;

        WorkspaceCenterLeftColumn.Width = new GridLength(1, GridUnitType.Star);
        WorkspaceCenterDividerColumn.Width = hasRightColumn ? new GridLength(1) : new GridLength(0);
        WorkspaceCenterRightColumn.Width = hasRightColumn ? new GridLength(1, GridUnitType.Star) : new GridLength(0);

        WorkspaceCenterTopRow.Height = new GridLength(1, GridUnitType.Star);
        WorkspaceCenterDividerRow.Height = hasBottomRow ? new GridLength(1) : new GridLength(0);
        WorkspaceCenterBottomRow.Height = hasBottomRow ? new GridLength(1, GridUnitType.Star) : new GridLength(0);

        WorkspaceVerticalDivider.Visibility = hasRightColumn ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceHorizontalDivider.Visibility = hasBottomRow ? Visibility.Visible : Visibility.Collapsed;

        WorkspacePane0Host.Visibility = Workspace0.HasTabs ? Visibility.Visible : Visibility.Collapsed;
        WorkspacePane1Host.Visibility = hasRightColumn && Workspace1.HasTabs ? Visibility.Visible : Visibility.Collapsed;
        WorkspacePane2Host.Visibility = hasBottomRow && Workspace2.HasTabs ? Visibility.Visible : Visibility.Collapsed;
        WorkspacePane3Host.Visibility = hasRightColumn && hasBottomRow && Workspace3.HasTabs ? Visibility.Visible : Visibility.Collapsed;
    }

    private void WorkspacePane_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: WorkspacePaneState pane })
        {
            return;
        }

        _selectedWorkspace = pane;
        if (pane.SelectedTab is not null)
        {
            SelectedTab = pane.SelectedTab;
        }
    }

    private void WorkspaceTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: WorkspacePaneState pane })
        {
            return;
        }

        if (pane.SelectedTab is not null)
        {
            _selectedWorkspace = pane;
            SelectedTab = pane.SelectedTab;
        }
    }

    private void WorkspaceTabControl_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not Selector selector || selector.DataContext is not WorkspacePaneState pane)
        {
            return;
        }

        if (pane.SelectedTab is null && pane.Tabs.Count > 0)
        {
            pane.SelectedTab = pane.Tabs[0];
        }

        if (pane.SelectedTab is not null && !ReferenceEquals(selector.SelectedItem, pane.SelectedTab))
        {
            selector.SelectedItem = pane.SelectedTab;
        }
    }

    private void TabHeader_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: ExplorerTabState tab })
        {
            _dragTabCandidate = tab;
            _tabDragStartPoint = e.GetPosition(this);
        }
    }

    private void TabHeader_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _dragTabCandidate is null || _tabDragStartPoint is null)
        {
            return;
        }

        var current = e.GetPosition(this);
        if (Math.Abs(current.X - _tabDragStartPoint.Value.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(current.Y - _tabDragStartPoint.Value.Y) < SystemParameters.MinimumVerticalDragDistance)
        {
            return;
        }

        var data = new DataObject(typeof(ExplorerTabState), _dragTabCandidate);
        DragDrop.DoDragDrop((DependencyObject)sender, data, DragDropEffects.Move);
        _dragTabCandidate = null;
        _tabDragStartPoint = null;
    }

    private void WorkspacePane_DragEnter(object sender, DragEventArgs e)
    {
        UpdateWorkspaceDropState(sender, e);
    }

    private void WorkspacePane_DragOver(object sender, DragEventArgs e)
    {
        UpdateWorkspaceDropState(sender, e);
        e.Handled = true;
    }

    private void WorkspacePane_DragLeave(object sender, DragEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: WorkspacePaneState pane })
        {
            pane.ClearDropState();
        }
    }

    private void WorkspacePane_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(typeof(ExplorerTabState)) ||
            e.Data.GetData(typeof(ExplorerTabState)) is not ExplorerTabState tab ||
            sender is not FrameworkElement { DataContext: WorkspacePaneState pane } element)
        {
            return;
        }

        var dropZone = GetDropZone(e.GetPosition(element), element.ActualWidth, element.ActualHeight);
        var destination = dropZone == WorkspaceDropZone.Center ? pane : GetAdjacentWorkspace(pane, dropZone) ?? pane;
        destination.IsVisible = true;

        MoveTabToWorkspace(tab, destination);
        ClearWorkspaceDropStates();
        e.Handled = true;
    }

    private void UpdateWorkspaceDropState(object sender, DragEventArgs e)
    {
        ClearWorkspaceDropStates();

        if (!e.Data.GetDataPresent(typeof(ExplorerTabState)) ||
            sender is not FrameworkElement { DataContext: WorkspacePaneState pane } element)
        {
            e.Effects = DragDropEffects.None;
            return;
        }

        var dropZone = GetDropZone(e.GetPosition(element), element.ActualWidth, element.ActualHeight);
        pane.IsDropActive = true;
        pane.DropHintText = GetDropHintText(dropZone);
        e.Effects = DragDropEffects.Move;
    }

    private void ClearWorkspaceDropStates()
    {
        foreach (var pane in _workspacePanes)
        {
            pane.ClearDropState();
        }
    }

    private static WorkspaceDropZone GetDropZone(Point position, double width, double height)
    {
        if (width <= 0 || height <= 0)
        {
            return WorkspaceDropZone.Center;
        }

        var normalizedX = position.X / width;
        var normalizedY = position.Y / height;

        if (normalizedX <= 0.22)
        {
            return WorkspaceDropZone.Left;
        }

        if (normalizedX >= 0.78)
        {
            return WorkspaceDropZone.Right;
        }

        if (normalizedY <= 0.22)
        {
            return WorkspaceDropZone.Top;
        }

        if (normalizedY >= 0.78)
        {
            return WorkspaceDropZone.Bottom;
        }

        return WorkspaceDropZone.Center;
    }

    private static string GetDropHintText(WorkspaceDropZone zone)
    {
        return zone switch
        {
            WorkspaceDropZone.Left => "左に並べて表示",
            WorkspaceDropZone.Right => "右に並べて表示",
            WorkspaceDropZone.Top => "上に並べて表示",
            WorkspaceDropZone.Bottom => "下に並べて表示",
            _ => "このペインへ移動"
        };
    }

    private WorkspacePaneState? GetAdjacentWorkspace(WorkspacePaneState pane, WorkspaceDropZone zone)
    {
        var destination = (pane.SlotIndex, zone) switch
        {
            (0, WorkspaceDropZone.Right) => Workspace1,
            (0, WorkspaceDropZone.Bottom) => Workspace2,
            (1, WorkspaceDropZone.Left) => Workspace0,
            (1, WorkspaceDropZone.Bottom) => Workspace3,
            (2, WorkspaceDropZone.Top) => Workspace0,
            (2, WorkspaceDropZone.Right) => Workspace3,
            (3, WorkspaceDropZone.Left) => Workspace2,
            (3, WorkspaceDropZone.Top) => Workspace1,
            _ => null
        };

        if (destination is not null)
        {
            destination.IsVisible = true;
        }

        return destination;
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        if (PresentationSource.FromVisual(this) is HwndSource source)
        {
            source.AddHook(WindowMessageHook);
        }
    }

    private IntPtr WindowMessageHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmMouseHWheel && HandleHorizontalScrollGesture(wParam))
        {
            handled = true;
        }

        return IntPtr.Zero;
    }

    private bool HandleHorizontalScrollGesture(IntPtr wParam)
    {
        var delta = unchecked((short)(((long)wParam >> 16) & 0xffff));
        if (delta == 0)
        {
            return false;
        }

        var scrollViewer = FindHorizontalScrollViewer(Mouse.DirectlyOver as DependencyObject);
        if (scrollViewer is null || scrollViewer.ScrollableWidth <= 0)
        {
            return false;
        }

        var nextOffset = delta > 0
            ? scrollViewer.HorizontalOffset - HorizontalScrollStep
            : scrollViewer.HorizontalOffset + HorizontalScrollStep;

        scrollViewer.ScrollToHorizontalOffset(Math.Clamp(nextOffset, 0, scrollViewer.ScrollableWidth));
        return true;
    }

    private static ScrollViewer? FindHorizontalScrollViewer(DependencyObject? start)
    {
        var scrollViewer = FindAncestor<ScrollViewer>(start);
        if (scrollViewer?.ScrollableWidth > 0)
        {
            return scrollViewer;
        }

        var dataGrid = FindAncestor<DataGrid>(start);
        scrollViewer = FindDescendant<ScrollViewer>(dataGrid);
        if (scrollViewer?.ScrollableWidth > 0)
        {
            return scrollViewer;
        }

        var treeView = FindAncestor<TreeView>(start);
        scrollViewer = FindDescendant<ScrollViewer>(treeView);
        return scrollViewer?.ScrollableWidth > 0 ? scrollViewer : null;
    }

    private void ApplyGrouping(ExplorerTabState tab)
    {
        tab.AssetsView.GroupDescriptions.Clear();

        var propertyName = tab.SelectedGroupingOption switch
        {
            "種類" => nameof(AssetItem.KindLabel),
            "更新日" => nameof(AssetItem.ModifiedGroupLabel),
            "フォルダー" => nameof(AssetItem.FolderGroupLabel),
            "お気に入り" => nameof(AssetItem.FavoriteGroupLabel),
            _ => null
        };

        if (propertyName is not null)
        {
            tab.AssetsView.GroupDescriptions.Add(new PropertyGroupDescription(propertyName));
        }

        tab.AssetsView.Refresh();
        tab.NotifyProperty(nameof(ExplorerTabState.ResultSummary));
    }

    private static ExplorerTabState? FindTabFromVisual(DependencyObject? start)
    {
        var current = start;
        while (current is not null)
        {
            if (current is FrameworkElement { DataContext: ExplorerTabState tab })
            {
                return tab;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private static T? FindAncestor<T>(DependencyObject? start) where T : DependencyObject
    {
        var current = start;
        while (current is not null)
        {
            if (current is T match)
            {
                return match;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private static T? FindDescendant<T>(DependencyObject? start) where T : DependencyObject
    {
        if (start is null)
        {
            return null;
        }

        for (var index = 0; index < VisualTreeHelper.GetChildrenCount(start); index++)
        {
            var child = VisualTreeHelper.GetChild(start, index);
            if (child is T match)
            {
                return match;
            }

            var nested = FindDescendant<T>(child);
            if (nested is not null)
            {
                return nested;
            }
        }

        return null;
    }

    private void AddLog(string message)
    {
        _activityLog.Insert(0, $"{DateTime.Now:HH:mm:ss}  {message}");
        while (_activityLog.Count > 50)
        {
            _activityLog.RemoveAt(_activityLog.Count - 1);
        }
    }

    private void ShowError(string message, Exception exception)
    {
        AddLog($"{message} {exception.Message}");
        MessageBox.Show($"{message}\n{exception.Message}", "MuseDock", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private static bool IsInteractiveElement(DependencyObject? origin)
    {
        var current = origin;
        while (current is not null)
        {
            if (current is Button or TextBox or ComboBox or DataGrid or ScrollBar or TabControl or TabItem or TreeView or TreeViewItem)
            {
                return true;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return false;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

internal sealed class PendingTransferOperation
{
    public required string SourceRootPath { get; init; }
    public required string[] SourcePaths { get; init; }
    public required bool IsMove { get; init; }
}

internal enum WorkspaceDropZone
{
    Center,
    Left,
    Right,
    Top,
    Bottom
}

public sealed class SubtractDoubleConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
        var source = value is double number ? number : 0d;
        var subtract = 0d;
        if (parameter is not null)
        {
            _ = double.TryParse(parameter.ToString(), out subtract);
        }

        return Math.Max(0d, source - subtract);
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}



