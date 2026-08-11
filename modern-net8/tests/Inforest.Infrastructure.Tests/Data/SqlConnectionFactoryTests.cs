using System.Data;
using Inforest.Infrastructure.Data;
using Inforest.Infrastructure.Exceptions;
using Microsoft.Extensions.Configuration;

namespace Inforest.Infrastructure.Tests.Data;

/// <summary>
/// Pruebas unitarias para <see cref="SqlConnectionFactory"/>.
/// Verifican la resolución de cadenas de conexión sin abrir BD real.
/// </summary>
public sealed class SqlConnectionFactoryTests
{
    // Cadena válida sintácticamente pero con servidor inexistente y timeout corto
    private const string InvalidServerConnStr =
        "Server=SERVIDOR_INEXISTENTE_12345;Database=INFOREST;" +
        "User Id=sa;Password=test;Connect Timeout=1;TrustServerCertificate=true;";

    private static IConfiguration BuildConfig(Dictionary<string, string?> values)
    {
        return new ConfigurationBuilder()
            .AddInMemoryCollection(values)
            .Build();
    }

    [Fact]
    public async Task CreateOpenConnectionAsync_MissingConnectionString_ThrowsInfrastructureException()
    {
        var config = BuildConfig(new Dictionary<string, string?>());
        var factory = new SqlConnectionFactory(config);

        var ex = await Assert.ThrowsAsync<InfrastructureException>(
            () => factory.CreateOpenConnectionAsync());

        Assert.Contains("Inforest", ex.Message);
    }

    [Fact]
    public async Task CreateOpenConnectionAsync_InvalidServer_ThrowsInfrastructureException()
    {
        var config = BuildConfig(new Dictionary<string, string?>
        {
            ["ConnectionStrings:Inforest"] = InvalidServerConnStr
        });
        var factory = new SqlConnectionFactory(config);

        var ex = await Assert.ThrowsAsync<InfrastructureException>(
            () => factory.CreateOpenConnectionAsync());

        Assert.Contains("conexión", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(ex.InnerException);
    }

    [Fact]
    public async Task CreateOpenConnectionAsync_AlternateDb_UsesDirectConnectionString()
    {
        // INFSEGURIDAD tiene cadena directa — debe intentar conectar a ese servidor (y fallar igual)
        var altConnStr =
            "Server=SERVIDOR_INEXISTENTE_12345;Database=INFSEGURIDAD;" +
            "User Id=sa;Password=test;Connect Timeout=1;TrustServerCertificate=true;";
        var config = BuildConfig(new Dictionary<string, string?>
        {
            ["ConnectionStrings:Inforest"] = InvalidServerConnStr,
            ["ConnectionStrings:INFSEGURIDAD"] = altConnStr
        });
        var factory = new SqlConnectionFactory(config);

        var ex = await Assert.ThrowsAsync<InfrastructureException>(
            () => factory.CreateOpenConnectionAsync("INFSEGURIDAD"));

        Assert.Contains("INFSEGURIDAD", ex.Message);
    }

    [Fact]
    public async Task CreateOpenConnectionAsync_FallbackToCatalogSubstitution_ThrowsOnInvalidServer()
    {
        // Sin cadena directa para ALMACEN, debe sustituir Initial Catalog en la cadena base
        var config = BuildConfig(new Dictionary<string, string?>
        {
            ["ConnectionStrings:Inforest"] = InvalidServerConnStr
        });
        var factory = new SqlConnectionFactory(config);

        var ex = await Assert.ThrowsAsync<InfrastructureException>(
            () => factory.CreateOpenConnectionAsync("ALMACEN"));

        Assert.Contains("ALMACEN", ex.Message);
    }
}
