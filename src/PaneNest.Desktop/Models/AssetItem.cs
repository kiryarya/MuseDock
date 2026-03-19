using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace PaneNest.Desktop.Models;

public sealed class AssetItem : INotifyPropertyChanged
{
    private string[] _tags = [];
    private string _note = string.Empty;
    private bool _isFavorite;

    public required string FilePath { get; init; }
    public required string Name { get; init; }
    public required string Extension { get; init; }
    public required string RelativeDirectory { get; init; }
    public required long Size { get; init; }
    public required DateTime LastModified { get; init; }
    public required AssetKind Kind { get; init; }
    public required bool IsHidden { get; init; }
    public required ImageSource? IconSource { get; init; }

    public string[] Tags
    {
        get => _tags;
        set
        {
            _tags = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(TagsDisplay));
        }
    }

    public string Note
    {
        get => _note;
        set
        {
            _note = value;
            OnPropertyChanged();
        }
    }

    public bool IsFavorite
    {
        get => _isFavorite;
        set
        {
            _isFavorite = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(FavoriteMark));
            OnPropertyChanged(nameof(FavoriteGroupLabel));
        }
    }

    public bool IsDirectory => Kind == AssetKind.Folder;

    public string KindLabel => Kind switch
    {
        AssetKind.Folder => "フォルダー",
        AssetKind.Image => "画像",
        AssetKind.Video => "動画",
        AssetKind.Text => "テキスト",
        AssetKind.Pdf => "PDF",
        _ => "その他"
    };

    public string RelativeDirectoryDisplay => string.IsNullOrWhiteSpace(RelativeDirectory) ? "このフォルダー" : RelativeDirectory;

    public string TagsDisplay => Tags.Length == 0 ? "タグなし" : string.Join(", ", Tags);

    public string SizeDisplay => IsDirectory ? "-" : MainWindowSizeFormatter.Format(Size);

    public string ModifiedDisplay => LastModified.ToString("yyyy-MM-dd HH:mm");

    public string FavoriteMark => IsFavorite ? "★" : string.Empty;

    public double RowOpacity => IsHidden ? 0.72 : 1.0;

    public string HiddenMark => IsHidden ? "隠し" : string.Empty;

    public string ModifiedGroupLabel
    {
        get
        {
            var today = DateTime.Today;
            if (LastModified.Date >= today)
            {
                return "今日";
            }

            if (LastModified.Date >= today.AddDays(-6))
            {
                return "過去 7 日";
            }

            if (LastModified.Date >= new DateTime(today.Year, today.Month, 1))
            {
                return "今月";
            }

            return LastModified.ToString("yyyy年M月");
        }
    }

    public string FolderGroupLabel => RelativeDirectoryDisplay;

    public string FavoriteGroupLabel => IsFavorite ? "お気に入り" : "通常";

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

internal static class MainWindowSizeFormatter
{
    public static string Format(long size)
    {
        string[] units = ["B", "KB", "MB", "GB", "TB"];
        double value = size;
        var unitIndex = 0;
        while (value >= 1024 && unitIndex < units.Length - 1)
        {
            value /= 1024;
            unitIndex += 1;
        }

        return $"{value:0.#} {units[unitIndex]}";
    }
}
