using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace MuseDock.Desktop;

internal sealed record ShortcutDefinition(string Key, string Label, string DefaultGesture);

internal sealed class ShortcutSettingsWindow : Window
{
    private readonly List<(ShortcutDefinition Definition, TextBox Editor)> _editors = [];

    public ShortcutSettingsWindow(
        IReadOnlyList<ShortcutDefinition> definitions,
        IReadOnlyDictionary<string, string> currentShortcuts)
    {
        Title = "ショートカット設定";
        Width = 560;
        Height = 480;
        MinWidth = 560;
        MinHeight = 480;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0C0F13"));
        Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F2F5F7"));

        var root = new Grid
        {
            Margin = new Thickness(18)
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var description = new StackPanel
        {
            Margin = new Thickness(0, 0, 0, 14)
        };
        description.Children.Add(new TextBlock
        {
            Text = "ショートカットを編集できます。",
            FontSize = 15,
            FontWeight = FontWeights.SemiBold
        });
        description.Children.Add(new TextBlock
        {
            Text = "形式例: Ctrl+T, Alt+Left, F2。空欄にすると無効です。",
            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#89A0B5")),
            Margin = new Thickness(0, 4, 0, 0)
        });
        root.Children.Add(description);

        var rowsPanel = new StackPanel();
        foreach (var definition in definitions)
        {
            var rowBorder = new Border
            {
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(12),
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#131920")),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#263545")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8)
            };

            var rowGrid = new Grid();
            rowGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rowGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rowGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            rowGrid.Children.Add(new TextBlock
            {
                Text = definition.Label,
                FontWeight = FontWeights.SemiBold
            });

            var defaultText = new TextBlock
            {
                Text = $"既定: {definition.DefaultGesture}",
                Margin = new Thickness(0, 4, 0, 8),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#89A0B5"))
            };
            Grid.SetRow(defaultText, 1);
            rowGrid.Children.Add(defaultText);

            var editor = new TextBox
            {
                Text = currentShortcuts.TryGetValue(definition.Key, out var shortcut) ? shortcut : definition.DefaultGesture
            };
            Grid.SetRow(editor, 2);
            rowGrid.Children.Add(editor);

            _editors.Add((definition, editor));
            rowBorder.Child = rowGrid;
            rowsPanel.Children.Add(rowBorder);
        }

        var scrollViewer = new ScrollViewer
        {
            Content = rowsPanel,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        Grid.SetRow(scrollViewer, 1);
        root.Children.Add(scrollViewer);

        var actions = new Grid
        {
            Margin = new Thickness(0, 14, 0, 0)
        };
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var restoreButton = new Button
        {
            Content = "既定に戻す",
            MinWidth = 92,
            Margin = new Thickness(0, 0, 8, 0)
        };
        restoreButton.Click += RestoreDefaults_Click;
        Grid.SetColumn(restoreButton, 1);
        actions.Children.Add(restoreButton);

        var cancelButton = new Button
        {
            Content = "キャンセル",
            MinWidth = 92,
            Margin = new Thickness(0, 0, 8, 0),
            IsCancel = true
        };
        Grid.SetColumn(cancelButton, 2);
        actions.Children.Add(cancelButton);

        var saveButton = new Button
        {
            Content = "保存",
            MinWidth = 92,
            IsDefault = true
        };
        saveButton.Click += Save_Click;
        Grid.SetColumn(saveButton, 3);
        actions.Children.Add(saveButton);

        Grid.SetRow(actions, 2);
        root.Children.Add(actions);
        Content = root;
    }

    public Dictionary<string, string> BuildShortcutMap()
    {
        return _editors.ToDictionary(
            item => item.Definition.Key,
            item => item.Editor.Text.Trim(),
            StringComparer.OrdinalIgnoreCase);
    }

    private void RestoreDefaults_Click(object? sender, RoutedEventArgs e)
    {
        foreach (var (definition, editor) in _editors)
        {
            editor.Text = definition.DefaultGesture;
        }
    }

    private void Save_Click(object? sender, RoutedEventArgs e)
    {
        var usedGestures = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (definition, editor) in _editors)
        {
            var gestureText = editor.Text.Trim();
            if (string.IsNullOrWhiteSpace(gestureText))
            {
                continue;
            }

            if (!TryParseKeyGesture(gestureText, out _))
            {
                MessageBox.Show(
                    $"{definition.Label} のショートカット形式が正しくありません。",
                    "ショートカット設定",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (usedGestures.TryGetValue(gestureText, out var existing))
            {
                MessageBox.Show(
                    $"{existing} と {definition.Label} に同じショートカットが設定されています。",
                    "ショートカット設定",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            usedGestures[gestureText] = definition.Label;
        }

        DialogResult = true;
        Close();
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
}

internal sealed class DelegateCommand(Action<object?> execute) : ICommand
{
    public event EventHandler? CanExecuteChanged;

    public bool CanExecute(object? parameter) => true;

    public void Execute(object? parameter) => execute(parameter);

    public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}
