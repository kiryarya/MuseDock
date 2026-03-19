using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace MuseDock.Desktop.Services;

internal static class FileIconService
{
    private const uint FileAttributeDirectory = 0x10;
    private const uint FileAttributeNormal = 0x80;
    private const uint ShgfiIcon = 0x100;
    private const uint ShgfiSmallIcon = 0x1;
    private const uint ShgfiUseFileAttributes = 0x10;

    private static readonly ConcurrentDictionary<string, ImageSource?> IconCache = new(StringComparer.OrdinalIgnoreCase);

    public static ImageSource? GetFolderIcon()
    {
        return IconCache.GetOrAdd("folder", _ => GetShellIcon("folder", FileAttributeDirectory));
    }

    public static ImageSource? GetFileIcon(string filePath, string extension)
    {
        var key = string.IsNullOrWhiteSpace(extension) ? "file" : extension;
        return IconCache.GetOrAdd(key, _ => GetShellIcon(string.IsNullOrWhiteSpace(extension) ? filePath : extension, FileAttributeNormal));
    }

    private static ImageSource? GetShellIcon(string pathOrExtension, uint fileAttributes)
    {
        try
        {
            var result = SHGetFileInfo(
                pathOrExtension,
                fileAttributes,
                out var shellInfo,
                (uint)Marshal.SizeOf<ShFileInfo>(),
                ShgfiIcon | ShgfiSmallIcon | ShgfiUseFileAttributes);

            if (result == IntPtr.Zero || shellInfo.IconHandle == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                var image = Imaging.CreateBitmapSourceFromHIcon(
                    shellInfo.IconHandle,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromWidthAndHeight(16, 16));
                image.Freeze();
                return image;
            }
            finally
            {
                DestroyIcon(shellInfo.IconHandle);
            }
        }
        catch
        {
            return null;
        }
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(
        string pszPath,
        uint dwFileAttributes,
        out ShFileInfo psfi,
        uint cbFileInfo,
        uint uFlags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ShFileInfo
    {
        public IntPtr IconHandle;
        public int IconIndex;
        public uint Attributes;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string DisplayName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string TypeName;
    }
}
