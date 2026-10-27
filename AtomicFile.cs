using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace SnapSlate;

internal static class AtomicFile
{
    public static void Write(string path, Action<string> writeTemporaryFile)
    {
        var temporary = CreateTemporaryPath(path);
        try
        {
            writeTemporaryFile(temporary);
            Commit(temporary, Path.GetFullPath(path));
        }
        finally
        {
            File.Delete(temporary);
        }
    }

    public static async Task WriteTextAsync(string path, string content)
    {
        var temporary = CreateTemporaryPath(path);
        try
        {
            await using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous))
            {
                await stream.WriteAsync(Encoding.UTF8.GetBytes(content));
                await stream.FlushAsync();
                stream.Flush(flushToDisk: true);
            }

            Commit(temporary, Path.GetFullPath(path));
        }
        finally
        {
            File.Delete(temporary);
        }
    }

    private static string CreateTemporaryPath(string path)
    {
        var absolute = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(absolute)!;
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, $".{Path.GetFileName(absolute)}.{Guid.NewGuid():N}.tmp");
    }

    private static void Commit(string temporary, string destination)
    {
        if (File.Exists(destination))
        {
            File.Replace(temporary, destination, destination + ".bak", ignoreMetadataErrors: true);
        }
        else
        {
            File.Move(temporary, destination);
        }
    }
}
