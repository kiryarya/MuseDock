using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace FilePilot.Desktop.Models;

public sealed class DirectoryTreeNode : INotifyPropertyChanged
{
    private bool _isExpanded;
    private bool _isSelected;
    private bool _areChildrenLoaded;

    public required string Path { get; init; }

    public required string Label { get; init; }

    public bool IsPlaceholder { get; init; }

    public ObservableCollection<DirectoryTreeNode> Children { get; } = [];

    public bool AreChildrenLoaded
    {
        get => _areChildrenLoaded;
        set => SetField(ref _areChildrenLoaded, value);
    }

    public bool IsExpanded
    {
        get => _isExpanded;
        set => SetField(ref _isExpanded, value);
    }

    public bool IsSelected
    {
        get => _isSelected;
        set => SetField(ref _isSelected, value);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public static DirectoryTreeNode CreatePlaceholder()
    {
        return new DirectoryTreeNode
        {
            Path = string.Empty,
            Label = string.Empty,
            IsPlaceholder = true
        };
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
