using System.Data;
using Inforest.Application.Interfaces;
using Inforest.Infrastructure.Exceptions;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace Inforest.Infrastructure.Data;

/// <summary>
/// Implementación de <see cref="IDbConnectionFactory"/> usando Microsoft.Data.SqlClient.
/// <para>
/// Legacy: <c>clsComando.cls → CreateCmdSp</c> recibía una <c>ADODB.Connection</c> ya abierta.
/// Esta fábrica centraliza la apertura y permite reuso seguro de la cadena de conexión
/// configurada en <c>appsettings.json</c> bajo la clave <c>ConnectionStrings:Inforest</c>.
/// </para>
/// Las bases de datos adicionales (INFSEGURIDAD, ALMACEN, CENTRALDELIVERY, FACTURACION)
/// se resuelven sustituyendo el <c>Initial Catalog</c> en la cadena base.
/// </summary>
internal sealed class SqlConnectionFactory : IDbConnectionFactory
{
    private const string DefaultKey = "Inforest";

    private readonly IConfiguration _configuration;

    public SqlConnectionFactory(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    /// <inheritdoc />
    public async Task<IDbConnection> CreateOpenConnectionAsync(CancellationToken cancellationToken = default)
    {
        return await CreateOpenConnectionAsync(DefaultKey, cancellationToken);
    }

    /// <inheritdoc />
    public async Task<IDbConnection> CreateOpenConnectionAsync(string databaseName, CancellationToken cancellationToken = default)
    {
        var connectionString = ResolveConnectionString(databaseName);
        var connection = new SqlConnection(connectionString);

        try
        {
            await connection.OpenAsync(cancellationToken);
            return connection;
        }
        catch (SqlException ex)
        {
            await connection.DisposeAsync();
            throw new InfrastructureException(
                $"No se pudo abrir la conexión a la base de datos '{databaseName}'. Verificar configuración y disponibilidad del servidor SQL.",
                ex);
        }
    }

    private string ResolveConnectionString(string databaseName)
    {
        // 1. Buscar clave exacta en ConnectionStrings (p. ej. "INFSEGURIDAD")
        var direct = _configuration.GetConnectionString(databaseName) + "User ID=infhotel;Password=4gust1n-fl0r14n;Max Pool Size=1000; ";
        if (!string.IsNullOrWhiteSpace(direct))
            return direct;

        // 2. Si no hay clave específica, tomar la cadena base y reemplazar Initial Catalog
        var baseString = _configuration.GetConnectionString(DefaultKey)
            ?? throw new InfrastructureException(
                $"No se encontró la cadena de conexión '{DefaultKey}' en la configuración. " +
                "Agregar 'ConnectionStrings:Inforest' en appsettings.json.");

        if (string.Equals(databaseName, DefaultKey, StringComparison.OrdinalIgnoreCase))
            return baseString;

        var builder = new SqlConnectionStringBuilder(baseString)
        {
            InitialCatalog = databaseName
        };
        return builder.ConnectionString;
    }
}
