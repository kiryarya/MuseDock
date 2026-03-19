using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Data;

namespace FilePilot.Desktop.Models;

public sealed class ExplorerTabState : INotifyPropertyChanged
{
    private string _title = "新しいタブ";
    private string _currentDrivePath = string.Empty;
    private string _currentLocationPath = string.Empty;
    private string _statusText = "準備完了";
    private string _searchText = string.Empty;
    private string _selectedTypeFilter = "すべて";
    private string _selectedGroupingOption = "なし";
    private string _pendingTagText = string.Empty;
    private bool _favoritesOnly;
    private NavigationItem? _selectedDrive;
    private NavigationItem? _selectedChildFolder;
    private AssetItem? _selectedAsset;
    private int _historyIndex = -1;
    private string? _previewImagePath;
    private Uri? _previewVideoUri;
    private string _previewText = string.Empty;
    private string _previewFallbackText = string.Empty;
    private bool _isImagePreviewVisible;
    private bool _isVideoPreviewVisible;
    private bool _isTextPreviewVisible;
    private bool _isFallbackPreviewVisible;

    public ExplorerTabState()
    {
        AssetsView = CollectionViewSource.GetDefaultView(Assets);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<AssetItem> Assets { get; } = [];

    public ObservableCollection<NavigationItem> ChildFolders { get; } = [];

    public ObservableCollection<string> TypeFilters { get; } =
    [
        "すべて",
        "フォルダー",
        "画像",
        "動画",
        "テキスト",
        "PDF",
        "その他"
    ];

    public ObservableCollection<string> GroupingOptions { get; } =
    [
        "なし",
        "種類",
        "更新日",
        "フォルダー",
        "お気に入り"
    ];

    public ObservableCollection<string> History { get; } = [];

    public ICollectionView AssetsView { get; }

    public LibraryMetadataDocument MetadataDocument { get; set; } = new();

    public CancellationTokenSource? LoadCts { get; set; }

    public bool SuppressNavigationSelection { get; set; }

    public string Title
    {
        get => _title;
        set => SetField(ref _title, value);
    }

    public string CurrentDrivePath
    {
        get => _currentDrivePath;
        set => SetField(ref _currentDrivePath, value);
    }

    public string CurrentLocationPath
    {
        get => _currentLocationPath;
        set
        {
            if (SetField(ref _currentLocationPath, value))
            {
                OnPropertyChanged(nameof(CurrentLocationName));
                OnPropertyChanged(nameof(CanGoUp));
            }
        }
    }

    public string CurrentLocationName =>
        string.IsNullOrWhiteSpace(CurrentLocationPath)
            ? "未選択"
            : Path.GetFileName(CurrentLocationPath.TrimEnd(Path.DirectorySeparatorChar)) switch
            {
                "" => CurrentLocationPath,
                var name => name
            };

    public string StatusText
    {
        get => _statusText;
        set => SetField(ref _statusText, value);
    }

    public string SearchText
    {
        get => _searchText;
        set => SetField(ref _searchText, value);
    }

    public string SelectedTypeFilter
    {
        get => _selectedTypeFilter;
        set => SetField(ref _selectedTypeFilter, value);
    }

    public string SelectedGroupingOption
    {
        get => _selectedGroupingOption;
        set => SetField(ref _selectedGroupingOption, value);
    }

    public string PendingTagText
    {
        get => _pendingTagText;
        set => SetField(ref _pendingTagText, value);
    }

    public bool FavoritesOnly
    {
        get => _favoritesOnly;
        set => SetField(ref _favoritesOnly, value);
    }

    public NavigationItem? SelectedDrive
    {
        get => _selectedDrive;
        set => SetField(ref _selectedDrive, value);
    }

    public NavigationItem? SelectedChildFolder
    {
        get => _selectedChildFolder;
        set => SetField(ref _selectedChildFolder, value);
    }

    public AssetItem? SelectedAsset
    {
        get => _selectedAsset;
        set
        {
            if (SetField(ref _selectedAsset, value))
            {
                OnPropertyChanged(nameof(HasSelection));
            }
        }
    }

    public int HistoryIndex
    {
        get => _historyIndex;
        set
        {
            if (SetField(ref _historyIndex, value))
            {
                OnPropertyChanged(nameof(CanGoBack));
                OnPropertyChanged(nameof(CanGoForward));
            }
        }
    }

    public bool CanGoUp => !string.IsNullOrWhiteSpace(CurrentLocationPath) &&
                           Directory.GetParent(CurrentLocationPath) is not null;

    public bool CanGoBack => HistoryIndex > 0;

    public bool CanGoForward => HistoryIndex >= 0 && HistoryIndex < History.Count - 1;

    public bool HasSelection => SelectedAsset is not null;

    public string ResultSummary =>
        Assets.Count == 0 ? "項目はありません" : $"{Assets.Count} 件を表示中";

    public string? PreviewImagePath
    {
        get => _previewImagePath;
        set => SetField(ref _previewImagePath, value);
    }

    public Uri? PreviewVideoUri
    {
        get => _previewVideoUri;
        set => SetField(ref _previewVideoUri, value);
    }

    public string PreviewText
    {
        get => _previewText;
        set => SetField(ref _previewText, value);
    }

    public string PreviewFallbackText
    {
        get => _previewFallbackText;
        set => SetField(ref _previewFallbackText, value);
    }

    public bool IsImagePreviewVisible
    {
        get => _isImagePreviewVisible;
        set => SetField(ref _isImagePreviewVisible, value);
    }

    public bool IsVideoPreviewVisible
    {
        get => _isVideoPreviewVisible;
        set => SetField(ref _isVideoPreviewVisible, value);
    }

    public bool IsTextPreviewVisible
    {
        get => _isTextPreviewVisible;
        set => SetField(ref _isTextPreviewVisible, value);
    }

    public bool IsFallbackPreviewVisible
    {
        get => _isFallbackPreviewVisible;
        set => SetField(ref _isFallbackPreviewVisible, value);
    }

    public void UpdateTitle()
    {
        Title = string.IsNullOrWhiteSpace(CurrentLocationPath)
            ? "新しいタブ"
            : CurrentLocationName;
    }

    public void RefreshDerivedState()
    {
        AssetsView.Refresh();
        NotifyProperty(nameof(CanGoUp));
        NotifyProperty(nameof(CanGoBack));
        NotifyProperty(nameof(CanGoForward));
        NotifyProperty(nameof(ResultSummary));
        NotifyProperty(nameof(HasSelection));
    }

    public void ResetPreview()
    {
        PreviewImagePath = null;
        PreviewVideoUri = null;
        PreviewText = string.Empty;
        PreviewFallbackText = string.Empty;
        IsImagePreviewVisible = false;
        IsVideoPreviewVisible = false;
        IsTextPreviewVisible = false;
        IsFallbackPreviewVisible = false;
    }

    public void NotifyProperty(string propertyName)
    {
        OnPropertyChanged(propertyName);
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
