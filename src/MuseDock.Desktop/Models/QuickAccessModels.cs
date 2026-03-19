using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MuseDock.Desktop.Models;

public sealed class QuickAccessFolder : INotifyPropertyChanged
{
    private string _path = string.Empty;
    private string _label = string.Empty;

    public event PropertyChangedEventHandler? PropertyChanged;

    public string Path
    {
        get => _path;
        set => SetField(ref _path, value);
    }

    public string Label
    {
        get => _label;
        set => SetField(ref _label, value);
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }
}

public sealed class QuickAccessGroup : INotifyPropertyChanged
{
    private string _name = string.Empty;

    public QuickAccessGroup()
    {
        Folders.CollectionChanged += Folders_CollectionChanged;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<QuickAccessFolder> Folders { get; } = [];

    public string Name
    {
        get => _name;
        set => SetField(ref _name, value);
    }

    public int FolderCount => Folders.Count;

    private void Folders_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FolderCount)));
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }
}

public sealed class QuickAccessDocument
{
    public List<QuickAccessGroupDocument> Groups { get; set; } = [];
}

public sealed class QuickAccessGroupDocument
{
    public string Name { get; set; } = string.Empty;

    public List<QuickAccessFolderDocument> Folders { get; set; } = [];
}

public sealed class QuickAccessFolderDocument
{
    public string Path { get; set; } = string.Empty;

    public string Label { get; set; } = string.Empty;
}
