using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace PaneNest.Desktop.Models;

public sealed class WorkspacePaneState : INotifyPropertyChanged
{
    private ExplorerTabState? _selectedTab;
    private bool _isVisible;
    private bool _isDropActive;
    private string _dropHintText = string.Empty;

    public WorkspacePaneState(int slotIndex)
    {
        SlotIndex = slotIndex;
        Tabs.CollectionChanged += Tabs_CollectionChanged;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public int SlotIndex { get; }

    public ObservableCollection<ExplorerTabState> Tabs { get; } = [];

    public ExplorerTabState? SelectedTab
    {
        get => _selectedTab;
        set
        {
            if (SetField(ref _selectedTab, value))
            {
                OnPropertyChanged(nameof(HasSelection));
            }
        }
    }

    public bool IsVisible
    {
        get => _isVisible;
        set => SetField(ref _isVisible, value);
    }

    public bool IsDropActive
    {
        get => _isDropActive;
        set => SetField(ref _isDropActive, value);
    }

    public string DropHintText
    {
        get => _dropHintText;
        set => SetField(ref _dropHintText, value);
    }

    public bool HasTabs => Tabs.Count > 0;

    public bool HasSelection => SelectedTab is not null;

    public void ClearDropState()
    {
        IsDropActive = false;
        DropHintText = string.Empty;
    }

    private void Tabs_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        OnPropertyChanged(nameof(HasTabs));

        if (_selectedTab is not null && !Tabs.Contains(_selectedTab))
        {
            SelectedTab = Tabs.FirstOrDefault();
        }
        else if (_selectedTab is null && Tabs.Count > 0)
        {
            SelectedTab = Tabs[0];
        }
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
