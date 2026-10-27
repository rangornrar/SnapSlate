using System;
using System.Diagnostics;
using System.IO;

namespace SnapSlate;

internal static class AppDiagnostics
{
    public static void Record(Exception exception)
    {
        // Log only error metadata: paths, image content and HTTP tokens must stay out of logs.
        var message = $"{DateTimeOffset.UtcNow:O} {exception.GetType().FullName} 0x{exception.HResult:X8}";
        Debug.WriteLine(message);
        try
        {
            var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SnapSlate");
            Directory.CreateDirectory(folder);
            File.AppendAllText(Path.Combine(folder, "errors.log"), message + Environment.NewLine);
        }
        catch (Exception loggingError)
        {
            Debug.WriteLine(loggingError.GetType().Name);
        }
    }
}
