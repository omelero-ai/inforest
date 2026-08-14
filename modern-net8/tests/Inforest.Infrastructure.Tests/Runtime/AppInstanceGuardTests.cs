using Inforest.Infrastructure.Runtime;

namespace Inforest.Infrastructure.Tests.Runtime;

public sealed class AppInstanceGuardTests
{
    [Fact]
    public void TryAcquire_PrimeraInstancia_RetornaTrue()
    {
        var sut = new AppInstanceGuard();
        var key = $"InfoRest-POS-{Guid.NewGuid():N}";

        var ok = sut.TryAcquire(key, out var lease);

        try
        {
            Assert.True(ok);
            Assert.NotNull(lease);
        }
        finally
        {
            lease.Dispose();
        }
    }

    [Fact]
    public void TryAcquire_SegundaInstanciaMientrasPrimeraActiva_RetornaFalse()
    {
        var key = $"InfoRest-POS-{Guid.NewGuid():N}";
        var guard1 = new AppInstanceGuard();
        var guard2 = new AppInstanceGuard();

        Assert.True(guard1.TryAcquire(key, out var lease1));
        try
        {
            var ok2 = guard2.TryAcquire(key, out var lease2);
            try
            {
                Assert.False(ok2);
            }
            finally
            {
                lease2.Dispose();
            }
        }
        finally
        {
            lease1.Dispose();
        }
    }

    [Fact]
    public void TryAcquire_LiberaMutexAlDispose_PermiteNuevaInstancia()
    {
        var key = $"InfoRest-POS-{Guid.NewGuid():N}";
        var guard = new AppInstanceGuard();

        Assert.True(guard.TryAcquire(key, out var lease1));
        lease1.Dispose();

        var ok2 = guard.TryAcquire(key, out var lease2);
        try
        {
            Assert.True(ok2);
        }
        finally
        {
            lease2.Dispose();
        }
    }
}
