using System.IO;
using System.Text;
using System.Windows.Threading;

namespace MuseDock.Desktop;

public partial class App : System.Windows.Application
{
    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MuseDock",
        "logs");

    protected override void OnStartup(System.Windows.StartupEventArgs e)
    {
        RegisterGlobalExceptionHandlers();

        try
        {
            base.OnStartup(e);
            ShutdownMode = System.Windows.ShutdownMode.OnMainWindowClose;

            var window = new MainWindow();
            MainWindow = window;
            window.Show();
        }
        catch (Exception exception)
        {
            HandleFatalException("アプリ起動中に致命的なエラーが発生しました。", exception);
        }
    }

    private void RegisterGlobalExceptionHandlers()
    {
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        WriteErrorLog("UI スレッドで未処理例外が発生しました。", e.Exception);
        System.Windows.MessageBox.Show(
            $"アプリ実行中にエラーが発生しました。\n\n{e.Exception.Message}\n\nログ: {GetLatestLogPath()}",
            "MuseDock",
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Error);
        e.Handled = true;
        Shutdown(-1);
    }

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
        {
            WriteErrorLog("バックグラウンドで未処理例外が発生しました。", exception);
        }
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        WriteErrorLog("監視されていないタスク例外が発生しました。", e.Exception);
        e.SetObserved();
    }

    private void HandleFatalException(string message, Exception exception)
    {
        WriteErrorLog(message, exception);
        System.Windows.MessageBox.Show(
            $"{message}\n\n{exception.Message}\n\nログ: {GetLatestLogPath()}",
            "MuseDock",
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Error);
        Shutdown(-1);
    }

    private static void WriteErrorLog(string headline, Exception exception)
    {
        Directory.CreateDirectory(LogDirectory);
        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var logPath = Path.Combine(LogDirectory, $"startup-{timestamp}.log");

        var builder = new StringBuilder();
        builder.AppendLine(headline);
        builder.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        builder.AppendLine();
        builder.AppendLine(exception.ToString());
        File.WriteAllText(logPath, builder.ToString(), Encoding.UTF8);
    }

    private static string GetLatestLogPath()
    {
        if (!Directory.Exists(LogDirectory))
        {
            return "ログ未作成";
        }

        return Directory.GetFiles(LogDirectory, "*.log")
            .OrderByDescending(path => path)
            .FirstOrDefault() ?? "ログ未作成";
    }
}
