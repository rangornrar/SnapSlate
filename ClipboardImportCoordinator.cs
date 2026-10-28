using System;
using System.Threading;
using System.Threading.Tasks;

namespace SnapSlate;

internal sealed class ClipboardImportCoordinator
{
    private long _revision;
    private long _processedRevision;
    private int _busy;

    public void NotifyChanged() => Interlocked.Increment(ref _revision);

    public async Task<int> ImportAsync(bool force, Func<Task<byte[]?>> read,
        Func<byte[], Task<ScreenshotDocument>> prepare, Action<ScreenshotDocument> insert, Func<bool> canInsert)
    {
        if (Interlocked.CompareExchange(ref _busy, 1, 0) != 0) return 0;
        var imported = 0;
        try
        {
            while (force || Interlocked.Read(ref _revision) != _processedRevision)
            {
                force = false;
                var revision = Interlocked.Read(ref _revision);
                try
                {
                    var bytes = await read();
                    if (revision != Interlocked.Read(ref _revision)) continue;
                    if (bytes is { Length: > 0 })
                    {
                        var document = await prepare(bytes);
                        if (revision != Interlocked.Read(ref _revision)) continue;
                        if (canInsert()) { insert(document); imported++; }
                    }
                    _processedRevision = revision;
                }
                catch { _processedRevision = revision; throw; }
            }
            return imported;
        }
        finally { Volatile.Write(ref _busy, 0); }
    }
}
