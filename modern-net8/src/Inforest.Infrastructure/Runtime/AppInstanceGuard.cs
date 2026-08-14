using Inforest.Application.Interfaces;
using System.Text;

namespace Inforest.Infrastructure.Runtime;

/// <summary>
/// Legacy: modPuntoVenta.bas App.PrevInstance + MsgBox + End (BR-POS-001).
/// </summary>
public sealed class AppInstanceGuard : IAppInstanceGuard
{
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, byte> LocalLeases = new();

    public bool TryAcquire(string instanceKey, out IDisposable lease)
    {
        if (string.IsNullOrWhiteSpace(instanceKey))
        {
            lease = NoopLease.Instance;
            return false;
        }

        var mutexName = BuildMutexName(instanceKey);
        if (!LocalLeases.TryAdd(mutexName, 0))
        {
            lease = NoopLease.Instance;
            return false;
        }

        var mutex = new Mutex(initiallyOwned: false, mutexName);

        try
        {
            if (!mutex.WaitOne(0, false))
            {
                LocalLeases.TryRemove(mutexName, out _);
                mutex.Dispose();
                lease = NoopLease.Instance;
                return false;
            }

            lease = new MutexLease(mutexName, mutex);
            return true;
        }
        catch (AbandonedMutexException)
        {
            lease = new MutexLease(mutexName, mutex);
            return true;
        }
        catch
        {
            LocalLeases.TryRemove(mutexName, out _);
            mutex.Dispose();
            throw;
        }
    }

    private static string BuildMutexName(string instanceKey)
    {
        var builder = new StringBuilder(instanceKey.Length);
        foreach (var ch in instanceKey.Trim())
            builder.Append(char.IsLetterOrDigit(ch) || ch is '.' or '-' or '_' ? ch : '_');

        return $"Inforest_{builder}";
    }

    private sealed class MutexLease(string mutexName, Mutex mutex) : IDisposable
    {
        private readonly string _mutexName = mutexName;
        private Mutex? _mutex = mutex;
        private bool _disposed;

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;
            try
            {
                _mutex?.ReleaseMutex();
            }
            catch (ApplicationException)
            {
                // Ignore if mutex ownership is already lost.
            }
            finally
            {
                LocalLeases.TryRemove(_mutexName, out _);
                _mutex?.Dispose();
                _mutex = null;
            }
        }
    }

    private sealed class NoopLease : IDisposable
    {
        public static readonly NoopLease Instance = new();
        private NoopLease() { }
        public void Dispose() { }
    }
}
