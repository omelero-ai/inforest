using System.Data;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Fábrica de conexiones a SQL Server.
/// Legacy: <c>clsComando.cls</c> recibía una <c>ADODB.Connection</c> activa como parámetro.
/// En .NET 8, la fábrica centraliza la creación y permite inyección y pruebas sin BD real.
/// </summary>
public interface IDbConnectionFactory
{
    /// <summary>
    /// Crea y abre una nueva conexión a la base de datos principal (INFOREST).
    /// </summary>
    Task<IDbConnection> CreateOpenConnectionAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Crea y abre una conexión a una base de datos específica del sistema
    /// (p. ej. INFSEGURIDAD, ALMACEN, CENTRALDELIVERY, FACTURACION).
    /// </summary>
    Task<IDbConnection> CreateOpenConnectionAsync(string databaseName, CancellationToken cancellationToken = default);
}
