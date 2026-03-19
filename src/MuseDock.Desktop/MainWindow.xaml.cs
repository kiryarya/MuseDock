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
using System.Windows.Threading;
using MuseDock.Desktop.Models;
using MuseDock.Desktop.Services;
using Microsoft.VisualBasic.FileIO;

namespace MuseDock.Desktop;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private const int WmMouseHWheel = 0x020E;
    private const double HorizontalScrollStep = 72d;
    private static readonly ShortcutDefinition[] ShortcutDefinitions =
    [
        new("new_tab", "新しいタブ", "Ctrl+T"),
        new("close_tab", "タブを閉じる", "Ctrl+W"),
        new("back", "戻る", "Alt+Left"),
        new("forward", "進む", "Alt+Right"),
        new("up", "上へ", "Alt+Up"),
        new("rename", "名前変更", "F2")
    ];
    private static readonly DependencyPropertyDescriptor? ColumnWidthDescriptor =
        DependencyPropertyDescriptor.FromProperty(DataGridColumn.WidthProperty, typeof(DataGridColumn));
    private readonly AssetLibraryService _libraryService = new();
    private readonly LibraryMetadataStore _metadataStore = new();
    private readonly PluginService _pluginService = new();
    private readonly AppSettingsStore _appSettingsStore = new();
    private readonly QuickAccessStore _quickAccessStore = new();
    private readonly ObservableCollection<NavigationItem> _drives = [];
    private readonly ObservableCollection<DirectoryTreeNode> _directoryRoots = [];
    private readonly ObservableCollection<QuickAccessGroup> _quickAccessGroups = [];
    private readonly ObservableCollection<ExplorerTabState> _tabs = [];
    private readonly WorkspacePaneState[] _workspacePanes;
    private readonly ObservableCollection<string> _activityLog = [];
    private readonly Dictionary<DataGrid, List<ColumnWidthSubscription>> _assetGridColumnSubscriptions = [];
    private readonly HashSet<DataGrid> _columnWidthRestoreInProgress = [];
    private readonly Dictionary<ExplorerTabState, CancellationTokenSource> _pendingMetadataSaveOperations = [];
    private readonly List<InputBinding> _dynamicShortcutBindings = [];
    private List<PluginCommandDefinition> _pluginCommands = [];
    private AppSettingsDocument _appSettingsDocument = new();
    private QuickAccessGroup? _selectedQuickAccessGroup;
    private ExplorerTabState? _selectedTab;
    private WorkspacePaneState? _selectedWorkspace;
    private DirectoryTreeNode? _selectedDirectoryNode;
    private PendingTransferOperation? _pendingTransfer;
    private bool _isWorkspaceDropOverlayVisible;
    private string _workspaceDropOverlayText = string.Empty;
    private string _workspaceDropOverlayKey = string.Empty;
    private bool _suppressTreeSelection;
    private Point? _tabDragStartPoint;
    private ExplorerTabState? _dragTabCandidate;

    public MainWindow()
    {
        _workspacePanes = Enumerable.Range(0, 9)
            .Select(index => new WorkspacePaneState(index) { IsVisible = index == 0 })
            .ToArray();

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

    public ObservableCollection<QuickAccessGroup> QuickAccessGroups => _quickAccessGroups;

    public ObservableCollection<string> ActivityLog => _activityLog;

    public WorkspacePaneState Workspace0 => _workspacePanes[0];

    public WorkspacePaneState Workspace1 => _workspacePanes[1];

    public WorkspacePaneState Workspace2 => _workspacePanes[2];

    public WorkspacePaneState Workspace3 => _workspacePanes[3];

    public WorkspacePaneState Workspace4 => _workspacePanes[4];

    public WorkspacePaneState Workspace5 => _workspacePanes[5];

    public WorkspacePaneState Workspace6 => _workspacePanes[6];

    public WorkspacePaneState Workspace7 => _workspacePanes[7];

    public WorkspacePaneState Workspace8 => _workspacePanes[8];

    public QuickAccessGroup? SelectedQuickAccessGroup
    {
        get => _selectedQuickAccessGroup;
        set
        {
            if (_selectedQuickAccessGroup == value)
            {
                return;
            }

            _selectedQuickAccessGroup = value;
            OnPropertyChanged();
        }
    }

    public bool IsWorkspaceDropOverlayVisible
    {
        get => _isWorkspaceDropOverlayVisible;
        set
        {
            if (_isWorkspaceDropOverlayVisible == value)
            {
                return;
            }

            _isWorkspaceDropOverlayVisible = value;
            OnPropertyChanged();
        }
    }

    public string WorkspaceDropOverlayText
    {
        get => _workspaceDropOverlayText;
        set
        {
            if (_workspaceDropOverlayText == value)
            {
                return;
            }

            _workspaceDropOverlayText = value;
            OnPropertyChanged();
        }
    }

    public string WorkspaceDropOverlayKey
    {
        get => _workspaceDropOverlayKey;
        set
        {
            if (_workspaceDropOverlayKey == value)
            {
                return;
            }

            _workspaceDropOverlayKey = value;
            OnPropertyChanged();
        }
    }

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

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        CaptureVisibleGridColumnWidths();
        _appSettingsStore.Save(_appSettingsDocument);
        _quickAccessStore.Save(BuildQuickAccessDocument());

        foreach (var tab in _tabs.ToArray())
        {
            CancelScheduledMetadataSave(tab);

            if (string.IsNullOrWhiteSpace(tab.CurrentDrivePath))
            {
                continue;
            }

            try
            {
                _metadataStore.Save(tab.CurrentDrivePath, tab.MetadataDocument);
            }
            catch
            {
                // Ignore shutdown persistence failures and allow the app to close.
            }
        }
    }

    private async Task InitializeExplorerAsync()
    {
        try
        {
            await LoadSettingsAsync();
            ReloadPlugins(logSummary: true);
            await LoadQuickAccessAsync();
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

    private async Task LoadSettingsAsync()
    {
        _appSettingsDocument = await _appSettingsStore.LoadAsync(CancellationToken.None);
        EnsureDefaultShortcuts(_appSettingsDocument);
        ApplyShortcutBindings();
    }

    private void EnsureDefaultShortcuts(AppSettingsDocument document)
    {
        foreach (var definition in ShortcutDefinitions)
        {
            if (!document.Shortcuts.ContainsKey(definition.Key))
            {
                document.Shortcuts[definition.Key] = definition.DefaultGesture;
            }
        }
    }

    private async Task PersistSettingsAsync()
    {
        await _appSettingsStore.SaveAsync(_appSettingsDocument, CancellationToken.None);
    }

    private void ApplyShortcutBindings()
    {
        foreach (var binding in _dynamicShortcutBindings)
        {
            InputBindings.Remove(binding);
        }

        _dynamicShortcutBindings.Clear();

        foreach (var definition in ShortcutDefinitions)
        {
            if (!_appSettingsDocument.Shortcuts.TryGetValue(definition.Key, out var gestureText) ||
                !TryParseKeyGesture(gestureText, out var gesture))
            {
                continue;
            }

            var binding = new KeyBinding(CreateShortcutCommand(definition.Key), gesture);
            InputBindings.Add(binding);
            _dynamicShortcutBindings.Add(binding);
        }
    }

    private ICommand CreateShortcutCommand(string shortcutKey)
    {
        return new DelegateCommand(_ =>
        {
            switch (shortcutKey)
            {
                case "new_tab":
                    NewTab_Click(this, new RoutedEventArgs());
                    break;
                case "close_tab":
                    CloseSelectedTabFromShortcut();
                    break;
                case "back":
                    Back_Click(this, new RoutedEventArgs());
                    break;
                case "forward":
                    Forward_Click(this, new RoutedEventArgs());
                    break;
                case "up":
                    Up_Click(this, new RoutedEventArgs());
                    break;
                case "rename":
                    RenameSelected_Click(this, new RoutedEventArgs());
                    break;
            }
        });
    }

    private async void CloseSelectedTabFromShortcut()
    {
        if (SelectedTab is not null)
        {
            await CloseTabAsync(SelectedTab);
        }
    }

    private static bool TryParseKeyGesture(string? gestureText, out KeyGesture gesture)
    {
        gesture = default!;
        if (string.IsNullOrWhiteSpace(gestureText))
        {
            return false;
        }

        try
        {
            if (new KeyGestureConverter().ConvertFromString(gestureText) is KeyGesture parsed &&
                parsed.Key != Key.None)
            {
                gesture = parsed;
                return true;
            }
        }
        catch
        {
        }

        return false;
    }

    private async Task LoadQuickAccessAsync()
    {
        var document = await _quickAccessStore.LoadAsync(CancellationToken.None);
        _quickAccessGroups.Clear();

        foreach (var groupDocument in document.Groups)
        {
            if (string.IsNullOrWhiteSpace(groupDocument.Name))
            {
                continue;
            }

            var group = new QuickAccessGroup
            {
                Name = groupDocument.Name
            };

            foreach (var folderDocument in groupDocument.Folders)
            {
                if (string.IsNullOrWhiteSpace(folderDocument.Path))
                {
                    continue;
                }

                group.Folders.Add(new QuickAccessFolder
                {
                    Path = folderDocument.Path,
                    Label = string.IsNullOrWhiteSpace(folderDocument.Label)
                        ? Path.GetFileName(folderDocument.Path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
                        : folderDocument.Label
                });
            }

            _quickAccessGroups.Add(group);
        }

        SelectedQuickAccessGroup = _quickAccessGroups.FirstOrDefault();
    }

    private async Task PersistQuickAccessAsync()
    {
        await _quickAccessStore.SaveAsync(BuildQuickAccessDocument(), CancellationToken.None);
    }

    private QuickAccessDocument BuildQuickAccessDocument()
    {
        return new QuickAccessDocument
        {
            Groups = _quickAccessGroups.Select(group => new QuickAccessGroupDocument
            {
                Name = group.Name,
                Folders = group.Folders.Select(folder => new QuickAccessFolderDocument
                {
                    Path = folder.Path,
                    Label = folder.Label
                }).ToList()
            }).ToList()
        };
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
        PersistVisibleGridColumnWidths(tab);
        await FlushMetadataSaveAsync(tab);
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

    private async void AddQuickAccessGroup_Click(object sender, RoutedEventArgs e)
    {
        var groupName = PromptForText("クイックアクセス", "新しいグループ名を入力してください。", "よく使う");
        if (string.IsNullOrWhiteSpace(groupName))
        {
            return;
        }

        if (_quickAccessGroups.Any(group => string.Equals(group.Name, groupName, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show("同じ名前のグループが既にあります。", "MuseDock", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var group = new QuickAccessGroup
        {
            Name = groupName.Trim()
        };

        _quickAccessGroups.Add(group);
        SelectedQuickAccessGroup = group;
        await PersistQuickAccessAsync();
        AddLog($"クイックアクセスグループ {group.Name} を追加しました。");
    }

    private async void AddCurrentFolderToQuickAccess_Click(object sender, RoutedEventArgs e)
    {
        var path = _selectedDirectoryNode?.Path;
        if (string.IsNullOrWhiteSpace(path))
        {
            path = SelectedTab?.CurrentLocationPath;
        }

        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
        {
            AddLog("登録できるフォルダーが選択されていません。");
            return;
        }

        var group = await EnsureQuickAccessGroupAsync();
        if (group is null)
        {
            return;
        }

        if (group.Folders.Any(folder => string.Equals(folder.Path, path, StringComparison.OrdinalIgnoreCase)))
        {
            AddLog($"{path} は既に {group.Name} に登録されています。");
            return;
        }

        group.Folders.Add(new QuickAccessFolder
        {
            Path = path,
            Label = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)) switch
            {
                "" => path,
                var label => label
            }
        });

        await PersistQuickAccessAsync();
        AddLog($"{path} を {group.Name} に登録しました。");
    }

    private async void OpenQuickAccessFolder_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: QuickAccessFolder folder } || !Directory.Exists(folder.Path))
        {
            return;
        }

        if (SelectedTab is not null)
        {
            await NavigateToPathAsync(SelectedTab, folder.Path);
        }
        else
        {
            await CreateAndOpenTabAsync(folder.Path);
        }
    }

    private async void RemoveQuickAccessFolder_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: QuickAccessFolder folder, CommandParameter: QuickAccessGroup group })
        {
            return;
        }

        group.Folders.Remove(folder);
        await PersistQuickAccessAsync();
        AddLog($"{folder.Label} を {group.Name} から外しました。");
    }

    private async void RenameQuickAccessGroup_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: QuickAccessGroup group })
        {
            return;
        }

        var nextName = PromptForText("クイックアクセス", "グループ名を入力してください。", group.Name);
        if (string.IsNullOrWhiteSpace(nextName))
        {
            return;
        }

        nextName = nextName.Trim();
        if (_quickAccessGroups.Any(existing => !ReferenceEquals(existing, group) && string.Equals(existing.Name, nextName, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show("同じ名前のグループが既にあります。", "MuseDock", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        group.Name = nextName;
        await PersistQuickAccessAsync();
        AddLog($"クイックアクセスグループ名を {nextName} に変更しました。");
    }

    private async void DeleteQuickAccessGroup_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: QuickAccessGroup group })
        {
            return;
        }

        if (MessageBox.Show(
                $"グループ {group.Name} を削除しますか？",
                "MuseDock",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
        {
            return;
        }

        _quickAccessGroups.Remove(group);
        if (ReferenceEquals(SelectedQuickAccessGroup, group))
        {
            SelectedQuickAccessGroup = _quickAccessGroups.FirstOrDefault();
        }

        await PersistQuickAccessAsync();
        AddLog($"クイックアクセスグループ {group.Name} を削除しました。");
    }

    private async Task<QuickAccessGroup?> EnsureQuickAccessGroupAsync()
    {
        if (SelectedQuickAccessGroup is not null)
        {
            return SelectedQuickAccessGroup;
        }

        var groupName = PromptForText("クイックアクセス", "登録先グループ名を入力してください。", "よく使う");
        if (string.IsNullOrWhiteSpace(groupName))
        {
            return null;
        }

        var existing = _quickAccessGroups.FirstOrDefault(group =>
            string.Equals(group.Name, groupName.Trim(), StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            SelectedQuickAccessGroup = existing;
            return existing;
        }

        var group = new QuickAccessGroup
        {
            Name = groupName.Trim()
        };

        _quickAccessGroups.Add(group);
        SelectedQuickAccessGroup = group;
        await PersistQuickAccessAsync();
        return group;
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

    private async void OpenSettings_Click(object sender, RoutedEventArgs e)
    {
        var window = new ShortcutSettingsWindow(ShortcutDefinitions, _appSettingsDocument.Shortcuts)
        {
            Owner = this
        };

        if (window.ShowDialog() != true)
        {
            return;
        }

        _appSettingsDocument.Shortcuts = window.BuildShortcutMap();
        EnsureDefaultShortcuts(_appSettingsDocument);
        ApplyShortcutBindings();
        await PersistSettingsAsync();
        AddLog("ショートカット設定を保存しました。");
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

        dataGrid.ContextMenu = CreateGridContextMenu(tab);
    }

    private void AssetGrid_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not DataGrid dataGrid || dataGrid.DataContext is not ExplorerTabState tab)
        {
            return;
        }

        RegisterAssetGridColumnTracking(dataGrid, tab);
        ApplyStoredColumnWidths(dataGrid, tab);
    }

    private void AssetGrid_Unloaded(object sender, RoutedEventArgs e)
    {
        if (sender is DataGrid dataGrid)
        {
            UnregisterAssetGridColumnTracking(dataGrid);
        }
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

    private async void RunPluginCommand_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: PluginMenuContext menuContext })
        {
            return;
        }

        await ExecutePluginCommandAsync(menuContext);
    }

    private void ReloadPlugins_Click(object sender, RoutedEventArgs e)
    {
        ReloadPlugins(logSummary: true);
    }

    private void OpenPluginFolder_Click(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(_pluginService.UserPluginDirectory);
        Process.Start(new ProcessStartInfo
        {
            FileName = _pluginService.UserPluginDirectory,
            UseShellExecute = true
        });

        AddLog("プラグインフォルダーを開きました。");
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
            await Dispatcher.InvokeAsync(() => ApplyStoredColumnWidthsToVisibleGrids(tab), DispatcherPriority.Loaded, cancellationToken);
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
        var nextName = PromptForText("名前変更", "新しい名前を入力してください。", asset.Name);
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
            await _metadataStore.SaveAsync(tab.CurrentDrivePath, tab.MetadataDocument, CancellationToken.None)
                .ConfigureAwait(false);
        }
    }

    private void RegisterAssetGridColumnTracking(DataGrid dataGrid, ExplorerTabState tab)
    {
        if (ColumnWidthDescriptor is null || _assetGridColumnSubscriptions.ContainsKey(dataGrid))
        {
            return;
        }

        var subscriptions = new List<ColumnWidthSubscription>(dataGrid.Columns.Count);
        foreach (var column in dataGrid.Columns)
        {
            EventHandler handler = (_, _) =>
            {
                if (_columnWidthRestoreInProgress.Contains(dataGrid))
                {
                    return;
                }

                PersistGridColumnWidths(tab, dataGrid);
                ScheduleMetadataSave(tab);
            };

            ColumnWidthDescriptor.AddValueChanged(column, handler);
            subscriptions.Add(new ColumnWidthSubscription(column, handler));
        }

        _assetGridColumnSubscriptions[dataGrid] = subscriptions;
    }

    private void UnregisterAssetGridColumnTracking(DataGrid dataGrid)
    {
        if (ColumnWidthDescriptor is null || !_assetGridColumnSubscriptions.Remove(dataGrid, out var subscriptions))
        {
            return;
        }

        foreach (var subscription in subscriptions)
        {
            ColumnWidthDescriptor.RemoveValueChanged(subscription.Column, subscription.Handler);
        }

        _columnWidthRestoreInProgress.Remove(dataGrid);
    }

    private void ApplyStoredColumnWidthsToVisibleGrids(ExplorerTabState tab)
    {
        foreach (var dataGrid in FindDescendants<DataGrid>(this).Where(grid => ReferenceEquals(grid.DataContext, tab)))
        {
            RegisterAssetGridColumnTracking(dataGrid, tab);
            ApplyStoredColumnWidths(dataGrid, tab);
        }
    }

    private void ApplyStoredColumnWidths(DataGrid dataGrid, ExplorerTabState tab)
    {
        if (tab.MetadataDocument.ColumnWidths.Count == 0)
        {
            return;
        }

        _columnWidthRestoreInProgress.Add(dataGrid);

        try
        {
            foreach (var column in dataGrid.Columns)
            {
                var key = GetColumnStorageKey(column);
                if (key is null ||
                    !tab.MetadataDocument.ColumnWidths.TryGetValue(key, out var width) ||
                    width <= 0)
                {
                    continue;
                }

                column.Width = new DataGridLength(Math.Max(width, column.MinWidth), DataGridLengthUnitType.Pixel);
            }
        }
        finally
        {
            _columnWidthRestoreInProgress.Remove(dataGrid);
        }
    }

    private void PersistVisibleGridColumnWidths(ExplorerTabState tab)
    {
        foreach (var dataGrid in FindDescendants<DataGrid>(this).Where(grid => ReferenceEquals(grid.DataContext, tab)))
        {
            PersistGridColumnWidths(tab, dataGrid);
        }
    }

    private void CaptureVisibleGridColumnWidths()
    {
        foreach (var dataGrid in FindDescendants<DataGrid>(this))
        {
            if (dataGrid.DataContext is ExplorerTabState tab)
            {
                PersistGridColumnWidths(tab, dataGrid);
            }
        }
    }

    private void PersistGridColumnWidths(ExplorerTabState tab, DataGrid dataGrid)
    {
        if (string.IsNullOrWhiteSpace(tab.CurrentDrivePath))
        {
            return;
        }

        var widths = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in dataGrid.Columns)
        {
            var key = GetColumnStorageKey(column);
            if (key is null)
            {
                continue;
            }

            var width = column.Width.IsAbsolute ? column.Width.DisplayValue : column.ActualWidth;
            if (double.IsNaN(width) || width <= 0)
            {
                continue;
            }

            widths[key] = Math.Round(width, 2);
        }

        if (widths.Count > 0)
        {
            tab.MetadataDocument.ColumnWidths = widths;
        }
    }

    private void ScheduleMetadataSave(ExplorerTabState tab)
    {
        if (string.IsNullOrWhiteSpace(tab.CurrentDrivePath))
        {
            return;
        }

        CancelScheduledMetadataSave(tab);

        var cancellation = new CancellationTokenSource();
        _pendingMetadataSaveOperations[tab] = cancellation;
        _ = PersistMetadataDeferredAsync(tab, cancellation);
    }

    private async Task PersistMetadataDeferredAsync(ExplorerTabState tab, CancellationTokenSource cancellation)
    {
        try
        {
            await Task.Delay(300, cancellation.Token);
            await PersistMetadataAsync(tab);
        }
        catch (OperationCanceledException)
        {
        }
        finally
        {
            if (_pendingMetadataSaveOperations.TryGetValue(tab, out var current) && ReferenceEquals(current, cancellation))
            {
                _pendingMetadataSaveOperations.Remove(tab);
            }

            cancellation.Dispose();
        }
    }

    private void CancelScheduledMetadataSave(ExplorerTabState tab)
    {
        if (_pendingMetadataSaveOperations.Remove(tab, out var cancellation))
        {
            cancellation.Cancel();
            cancellation.Dispose();
        }
    }

    private async Task FlushMetadataSaveAsync(ExplorerTabState tab)
    {
        CancelScheduledMetadataSave(tab);
        await PersistMetadataAsync(tab);
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

    private ContextMenu CreateGridContextMenu(ExplorerTabState tab)
    {
        var menu = new ContextMenu();
        var newTabItem = new MenuItem { Header = "新しいタブ" };
        newTabItem.Click += NewTabFromContext_Click;
        menu.Items.Add(newTabItem);

        var pasteItem = new MenuItem { Header = "ここに貼り付け", IsEnabled = HasPendingTransfer };
        pasteItem.Click += PasteIntoCurrentDirectory_Click;
        menu.Items.Add(pasteItem);

        menu.Items.Add(new Separator());
        menu.Items.Add(CreatePluginMenu(tab, null));
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
        menu.Items.Add(CreatePluginMenu(tab, asset));
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

    private MenuItem CreatePluginMenu(ExplorerTabState tab, AssetItem? asset)
    {
        var pluginMenu = new MenuItem { Header = "プラグイン" };
        var invocationContext = CreatePluginInvocationContext(tab, asset);
        var matchingCommands = _pluginCommands
            .Where(command => command.Matches(invocationContext))
            .OrderBy(command => command.PluginName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(command => command.CommandName, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (matchingCommands.Length == 0)
        {
            pluginMenu.Items.Add(new MenuItem
            {
                Header = "利用できるプラグインはありません",
                IsEnabled = false
            });
        }
        else
        {
            foreach (var pluginGroup in matchingCommands.GroupBy(command => command.PluginId))
            {
                var first = pluginGroup.First();
                var pluginItem = new MenuItem
                {
                    Header = first.PluginName,
                    ToolTip = string.IsNullOrWhiteSpace(first.PluginDescription) ? null : first.PluginDescription
                };

                foreach (var command in pluginGroup)
                {
                    var commandItem = new MenuItem
                    {
                        Header = command.CommandName,
                        Tag = new PluginMenuContext(tab, command, invocationContext),
                        ToolTip = string.IsNullOrWhiteSpace(command.CommandDescription) ? null : command.CommandDescription
                    };
                    commandItem.Click += RunPluginCommand_Click;
                    pluginItem.Items.Add(commandItem);
                }

                pluginMenu.Items.Add(pluginItem);
            }
        }

        pluginMenu.Items.Add(new Separator());

        var reloadItem = new MenuItem { Header = "プラグインを再読み込み" };
        reloadItem.Click += ReloadPlugins_Click;
        pluginMenu.Items.Add(reloadItem);

        var openFolderItem = new MenuItem { Header = "プラグインフォルダーを開く" };
        openFolderItem.Click += OpenPluginFolder_Click;
        pluginMenu.Items.Add(openFolderItem);
        return pluginMenu;
    }

    private PluginInvocationContext CreatePluginInvocationContext(ExplorerTabState tab, AssetItem? asset)
    {
        var selectedItems = new List<PluginSelectedItem>();
        if (asset is not null)
        {
            selectedItems.Add(new PluginSelectedItem
            {
                FilePath = asset.FilePath,
                Name = asset.Name,
                Extension = asset.Extension,
                AssetKind = ToPluginAssetKind(asset.Kind),
                IsDirectory = asset.IsDirectory,
                Tags = [.. asset.Tags],
                Note = asset.Note,
                IsFavorite = asset.IsFavorite
            });
        }

        return new PluginInvocationContext
        {
            CurrentDirectory = tab.CurrentLocationPath,
            CurrentDrivePath = tab.CurrentDrivePath,
            SelectedItems = selectedItems
        };
    }

    private async Task ExecutePluginCommandAsync(PluginMenuContext menuContext)
    {
        try
        {
            var result = await _pluginService.ExecuteAsync(menuContext.Command, menuContext.InvocationContext, CancellationToken.None);
            AddLog($"プラグイン {menuContext.Command.PluginName} / {menuContext.Command.CommandName} を実行しました。");

            if (menuContext.Command.RefreshAfterRun)
            {
                await RefreshAffectedTabsAsync([menuContext.InvocationContext.CurrentDirectory], menuContext.InvocationContext.PrimarySelectedPath);
            }

            AddLog($"プラグインコンテキスト: {result.ContextPath}");
        }
        catch (Exception exception)
        {
            ShowError("プラグインの実行に失敗しました。", exception);
        }
    }

    private void ReloadPlugins(bool logSummary)
    {
        var catalog = _pluginService.LoadCatalog();
        _pluginCommands = catalog.Commands;

        if (logSummary)
        {
            AddLog($"プラグインを読み込みました。{_pluginCommands.Count} コマンド。");
        }

        foreach (var error in catalog.Errors)
        {
            AddLog($"プラグイン読み込みエラー: {error}");
        }
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
        foreach (var pane in _workspacePanes)
        {
            pane.IsVisible = pane.HasTabs || pane.SlotIndex == 0;
            if (!pane.HasTabs)
            {
                pane.SelectedTab = null;
            }
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

        var columnHasTabs = Enumerable.Range(0, 3)
            .Select(column => _workspacePanes.Any(pane => pane.HasTabs && GetPaneColumn(pane.SlotIndex) == column))
            .ToArray();
        var rowHasTabs = Enumerable.Range(0, 3)
            .Select(row => _workspacePanes.Any(pane => pane.HasTabs && GetPaneRow(pane.SlotIndex) == row))
            .ToArray();

        WorkspaceColumn0.Width = columnHasTabs[0] ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
        WorkspaceColumn1.Width = columnHasTabs[1] ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
        WorkspaceColumn2.Width = columnHasTabs[2] ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
        WorkspaceDividerColumn01.Width = HasVisibleColumnsAround(columnHasTabs, 0) ? new GridLength(1) : new GridLength(0);
        WorkspaceDividerColumn12.Width = HasVisibleColumnsAround(columnHasTabs, 1) ? new GridLength(1) : new GridLength(0);

        WorkspaceRow0.Height = rowHasTabs[0] ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
        WorkspaceRow1.Height = rowHasTabs[1] ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
        WorkspaceRow2.Height = rowHasTabs[2] ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
        WorkspaceDividerRow01.Height = HasVisibleRowsAround(rowHasTabs, 0) ? new GridLength(1) : new GridLength(0);
        WorkspaceDividerRow12.Height = HasVisibleRowsAround(rowHasTabs, 1) ? new GridLength(1) : new GridLength(0);

        WorkspaceVerticalDivider01.Visibility = WorkspaceDividerColumn01.Width.Value > 0 ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceVerticalDivider12.Visibility = WorkspaceDividerColumn12.Width.Value > 0 ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceHorizontalDivider01.Visibility = WorkspaceDividerRow01.Height.Value > 0 ? Visibility.Visible : Visibility.Collapsed;
        WorkspaceHorizontalDivider12.Visibility = WorkspaceDividerRow12.Height.Value > 0 ? Visibility.Visible : Visibility.Collapsed;

        for (var index = 0; index < _workspacePanes.Length; index++)
        {
            GetWorkspaceHost(index).Visibility = _workspacePanes[index].HasTabs ? Visibility.Visible : Visibility.Collapsed;
        }
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
        try
        {
            DragDrop.DoDragDrop((DependencyObject)sender, data, DragDropEffects.Move);
        }
        finally
        {
            ClearWorkspaceDropStates();
            _dragTabCandidate = null;
            _tabDragStartPoint = null;
        }
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
        ClearWorkspaceDropStates();
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
        pane.DropZoneKey = dropZone.ToString();
        IsWorkspaceDropOverlayVisible = true;
        WorkspaceDropOverlayText = pane.DropHintText;
        WorkspaceDropOverlayKey = pane.DropZoneKey;
        e.Effects = DragDropEffects.Move;
    }

    private void ClearWorkspaceDropStates()
    {
        foreach (var pane in _workspacePanes)
        {
            pane.ClearDropState();
        }

        IsWorkspaceDropOverlayVisible = false;
        WorkspaceDropOverlayText = string.Empty;
        WorkspaceDropOverlayKey = string.Empty;
    }

    private static WorkspaceDropZone GetDropZone(Point position, double width, double height)
    {
        if (width <= 0 || height <= 0)
        {
            return WorkspaceDropZone.Center;
        }

        var normalizedX = position.X / width;
        var normalizedY = position.Y / height;
        const double edgeThreshold = 0.24;

        var closest = new (WorkspaceDropZone Zone, double Distance)[]
        {
            (WorkspaceDropZone.Left, normalizedX),
            (WorkspaceDropZone.Right, 1d - normalizedX),
            (WorkspaceDropZone.Top, normalizedY),
            (WorkspaceDropZone.Bottom, 1d - normalizedY)
        }
            .OrderBy(item => item.Distance)
            .First();

        return closest.Distance <= edgeThreshold ? closest.Zone : WorkspaceDropZone.Center;
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
        var row = GetPaneRow(pane.SlotIndex);
        var column = GetPaneColumn(pane.SlotIndex);
        var destination = zone switch
        {
            WorkspaceDropZone.Left when column > 0 => GetWorkspaceAt(row, column - 1),
            WorkspaceDropZone.Right when column < 2 => GetWorkspaceAt(row, column + 1),
            WorkspaceDropZone.Top when row > 0 => GetWorkspaceAt(row - 1, column),
            WorkspaceDropZone.Bottom when row < 2 => GetWorkspaceAt(row + 1, column),
            _ => null
        };

        if (destination is not null)
        {
            destination.IsVisible = true;
        }

        return destination;
    }

    private WorkspacePaneState? GetWorkspaceAt(int row, int column)
    {
        var slotIndex = row * 3 + column;
        return slotIndex >= 0 && slotIndex < _workspacePanes.Length ? _workspacePanes[slotIndex] : null;
    }

    private static int GetPaneRow(int slotIndex)
    {
        return slotIndex / 3;
    }

    private static int GetPaneColumn(int slotIndex)
    {
        return slotIndex % 3;
    }

    private bool HasVisibleColumnsAround(IReadOnlyList<bool> visibility, int dividerIndex)
    {
        var hasLeft = visibility.Take(dividerIndex + 1).Any(isVisible => isVisible);
        var hasRight = visibility.Skip(dividerIndex + 1).Any(isVisible => isVisible);
        return hasLeft && hasRight;
    }

    private bool HasVisibleRowsAround(IReadOnlyList<bool> visibility, int dividerIndex)
    {
        var hasTop = visibility.Take(dividerIndex + 1).Any(isVisible => isVisible);
        var hasBottom = visibility.Skip(dividerIndex + 1).Any(isVisible => isVisible);
        return hasTop && hasBottom;
    }

    private ContentControl GetWorkspaceHost(int slotIndex)
    {
        return slotIndex switch
        {
            0 => WorkspacePane0Host,
            1 => WorkspacePane1Host,
            2 => WorkspacePane2Host,
            3 => WorkspacePane3Host,
            4 => WorkspacePane4Host,
            5 => WorkspacePane5Host,
            6 => WorkspacePane6Host,
            7 => WorkspacePane7Host,
            8 => WorkspacePane8Host,
            _ => throw new ArgumentOutOfRangeException(nameof(slotIndex))
        };
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

    private static IEnumerable<T> FindDescendants<T>(DependencyObject? start) where T : DependencyObject
    {
        if (start is null)
        {
            yield break;
        }

        for (var index = 0; index < VisualTreeHelper.GetChildrenCount(start); index++)
        {
            var child = VisualTreeHelper.GetChild(start, index);
            if (child is T match)
            {
                yield return match;
            }

            foreach (var nested in FindDescendants<T>(child))
            {
                yield return nested;
            }
        }
    }

    private static string? GetColumnStorageKey(DataGridColumn column)
    {
        return column.Header as string;
    }

    private static string ToPluginAssetKind(AssetKind kind)
    {
        return kind switch
        {
            AssetKind.Folder => "folder",
            AssetKind.Image => "image",
            AssetKind.Video => "video",
            AssetKind.Text => "text",
            AssetKind.Pdf => "pdf",
            _ => "other"
        };
    }

    private string? PromptForText(string title, string prompt, string initialValue)
    {
        var dialog = new TextPromptWindow(title, prompt, initialValue)
        {
            Owner = this
        };

        return dialog.ShowDialog() == true ? dialog.ResponseText : null;
    }

    private void AddLog(string message)
    {
        _activityLog.Insert(0, $"{DateTime.Now:HH:mm:ss}  {message}");
        while (_activityLog.Count > 50)
        {
            _activityLog.RemoveAt(_activityLog.Count - 1);
        }
    }

    public async Task<LayoutSelfTestResult> RunLayoutSelfTestAsync()
    {
        if (SelectedTab is null)
        {
            throw new InvalidOperationException("SelectedTab is not initialized.");
        }

        var testPath = GetLayoutSelfTestPath();
        if (!string.IsNullOrWhiteSpace(testPath) &&
            !string.Equals(SelectedTab.CurrentLocationPath, testPath, StringComparison.OrdinalIgnoreCase))
        {
            await NavigateToPathAsync(SelectedTab, testPath);
        }

        var secondTab = await CreateAndOpenTabAsync(SelectedTab.CurrentLocationPath);
        MoveTabToWorkspace(secondTab, Workspace3);
        ClearWorkspaceDropStates();

        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
        await Task.Delay(300);
        UpdateLayout();
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);

        var visibleGrids = FindDescendants<DataGrid>(this)
            .Where(grid => grid.Visibility == Visibility.Visible && grid.ActualHeight > 80 && grid.ActualWidth > 100)
            .Select(grid =>
            {
                var scrollViewer = FindDescendant<ScrollViewer>(grid);
                var visibleVerticalScrollBars = FindDescendants<ScrollBar>(grid)
                    .Count(scrollBar => scrollBar.Visibility == Visibility.Visible && scrollBar.Orientation == Orientation.Vertical);
                return new LayoutSelfTestGridInfo
                {
                    Height = Math.Round(grid.ActualHeight, 1),
                    Width = Math.Round(grid.ActualWidth, 1),
                    VerticalScrollVisibility = scrollViewer?.ComputedVerticalScrollBarVisibility.ToString() ?? "Unknown",
                    ScrollableHeight = scrollViewer is null ? 0 : Math.Round(scrollViewer.ScrollableHeight, 1),
                    VisibleVerticalScrollBarCount = visibleVerticalScrollBars
                };
            })
            .ToList();

        return new LayoutSelfTestResult
        {
            VisibleGridCount = visibleGrids.Count,
            Row0Height = Math.Round(WorkspaceRow0.ActualHeight, 1),
            Row1Height = Math.Round(WorkspaceRow1.ActualHeight, 1),
            Row2Height = Math.Round(WorkspaceRow2.ActualHeight, 1),
            Host0Visibility = WorkspacePane0Host.Visibility.ToString(),
            Host3Visibility = WorkspacePane3Host.Visibility.ToString(),
            OverlayVisible = IsWorkspaceDropOverlayVisible,
            Grids = visibleGrids
        };
    }

    private static string GetLayoutSelfTestPath()
    {
        var windowsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrWhiteSpace(windowsDirectory))
        {
            var system32 = Path.Combine(windowsDirectory, "System32");
            if (Directory.Exists(system32))
            {
                return system32;
            }

            if (Directory.Exists(windowsDirectory))
            {
                return windowsDirectory;
            }
        }

        return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
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

internal sealed class TextPromptWindow : Window
{
    private readonly TextBox _textBox;

    public TextPromptWindow(string title, string prompt, string initialValue)
    {
        Title = title;
        Width = 420;
        Height = 176;
        MinWidth = 420;
        MaxWidth = 420;
        ResizeMode = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        WindowStyle = WindowStyle.ToolWindow;
        Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#131920"));
        Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F2F5F7"));
        ShowInTaskbar = false;

        var root = new Grid
        {
            Margin = new Thickness(16)
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        root.Children.Add(new TextBlock
        {
            Text = prompt,
            FontSize = 13,
            Margin = new Thickness(0, 0, 0, 10),
            TextWrapping = TextWrapping.Wrap
        });

        _textBox = new TextBox
        {
            Text = initialValue,
            Margin = new Thickness(0, 0, 0, 14)
        };
        Grid.SetRow(_textBox, 1);
        root.Children.Add(_textBox);

        var buttons = new Grid();
        buttons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        buttons.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        buttons.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        Grid.SetRow(buttons, 3);

        var cancelButton = new Button
        {
            Content = "キャンセル",
            MinWidth = 88,
            Margin = new Thickness(0, 0, 8, 0),
            IsCancel = true
        };
        Grid.SetColumn(cancelButton, 1);
        buttons.Children.Add(cancelButton);

        var okButton = new Button
        {
            Content = "OK",
            MinWidth = 88,
            IsDefault = true
        };
        okButton.Click += (_, _) =>
        {
            DialogResult = true;
            Close();
        };
        Grid.SetColumn(okButton, 2);
        buttons.Children.Add(okButton);

        root.Children.Add(buttons);
        Content = root;

        Loaded += (_, _) =>
        {
            _textBox.Focus();
            _textBox.SelectAll();
        };
    }

    public string ResponseText => _textBox.Text.Trim();
}

internal sealed class ColumnWidthSubscription(DataGridColumn column, EventHandler handler)
{
    public DataGridColumn Column { get; } = column;
    public EventHandler Handler { get; } = handler;
}

internal sealed record PluginMenuContext(
    ExplorerTabState Tab,
    PluginCommandDefinition Command,
    PluginInvocationContext InvocationContext);

public sealed class LayoutSelfTestResult
{
    public int VisibleGridCount { get; init; }
    public double Row0Height { get; init; }
    public double Row1Height { get; init; }
    public double Row2Height { get; init; }
    public string Host0Visibility { get; init; } = string.Empty;
    public string Host3Visibility { get; init; } = string.Empty;
    public bool OverlayVisible { get; init; }
    public List<LayoutSelfTestGridInfo> Grids { get; init; } = [];
}

public sealed class LayoutSelfTestGridInfo
{
    public double Height { get; init; }
    public double Width { get; init; }
    public string VerticalScrollVisibility { get; init; } = string.Empty;
    public double ScrollableHeight { get; init; }
    public int VisibleVerticalScrollBarCount { get; init; }
}



