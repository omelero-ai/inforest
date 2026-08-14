using Microsoft.Extensions.DependencyInjection;

namespace Inforest.Application.Tests;

/// <summary>
/// Pruebas baseline de P3-01 — valida que el DI de Application se registra correctamente.
/// </summary>
public class DependencyInjectionTests
{
    [Fact]
    public void AddApplication_NoLanzaException()
    {
        var services = new ServiceCollection();
        var ex = Record.Exception(() => services.AddApplication());
        Assert.Null(ex);
    }

    [Fact]
    public void AddApplication_RetornaIServiceCollection()
    {
        var services = new ServiceCollection();
        var result = services.AddApplication();
        Assert.Same(services, result);
    }
}