using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Infrastructure.Delivery;

/// <summary>
/// Repositorio Dapper para locales del restaurante (multi-local).
/// <para>
/// Legacy: tabla <c>TLOCAL</c>.
/// </para>
/// Regla BR-DEL-007.
/// </summary>
internal sealed class LocalRepository : ILocalRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public LocalRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    /// <inheritdoc />
    public async Task<LocalRestaurante?> ObtenerPorCodigoAsync(string codigoLocal, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoLocal AS CodigoLocal, tDetallado AS Detallado,
                   tResumido AS Resumido, tcodigoSector AS CodigoSector,
                   tIP AS DireccionIp, tBaseDatosINF AS BaseDatosInf,
                   tBaseDatosALM AS BaseDatosAlm,
                   ISNULL(nMeta,0) AS Meta, ISNULL(lActivo,0) AS Activo,
                   ISNULL(lReplica,0) AS ParticipanEnReplica
            FROM TLOCAL
            WHERE tCodigoLocal = @CodigoLocal
            """;
        return await conn.QueryFirstOrDefaultAsync<LocalRestaurante>(sql, new { CodigoLocal = codigoLocal });
    }

    /// <inheritdoc />
    public async Task<IEnumerable<LocalRestaurante>> ObtenerActivosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoLocal AS CodigoLocal, tDetallado AS Detallado,
                   tResumido AS Resumido, tIP AS DireccionIp,
                   tBaseDatosINF AS BaseDatosInf, ISNULL(lActivo,0) AS Activo,
                   ISNULL(lReplica,0) AS ParticipanEnReplica
            FROM TLOCAL
            WHERE ISNULL(lActivo,0) = 1
            ORDER BY tCodigoLocal
            """;
        return await conn.QueryAsync<LocalRestaurante>(sql);
    }

    /// <inheritdoc />
    public async Task<IEnumerable<LocalRestaurante>> ObtenerConReplicaAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoLocal AS CodigoLocal, tDetallado AS Detallado,
                   tIP AS DireccionIp, tBaseDatosINF AS BaseDatosInf,
                   ISNULL(lActivo,0) AS Activo,
                   ISNULL(lReplica,0) AS ParticipanEnReplica
            FROM TLOCAL
            WHERE ISNULL(lReplica,0) = 1 AND ISNULL(lActivo,0) = 1
            ORDER BY tCodigoLocal
            """;
        return await conn.QueryAsync<LocalRestaurante>(sql);
    }
}
