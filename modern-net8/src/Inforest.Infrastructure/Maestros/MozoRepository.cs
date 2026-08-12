using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio Dapper de mozos.
/// Legacy: TMOZO / frmMozo.frm. Database: Inforest.
/// </summary>
internal sealed class MozoRepository : IMozoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public MozoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<Mozo>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoMozo AS CodigoMozo,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro
            FROM TMOZO
            ORDER BY tDetallado
            """;
        return (await conn.QueryAsync<Mozo>(sql)).AsList();
    }

    public async Task<Mozo?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoMozo AS CodigoMozo,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro
            FROM TMOZO
            WHERE tCodigoMozo = @Codigo
            """;
        return await conn.QueryFirstOrDefaultAsync<Mozo>(sql, new { Codigo = codigo });
    }

    public async Task<bool> InsertarAsync(Mozo mozo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO TMOZO (tCodigoMozo, tDetallado, tResumido, lActivo, tUsuario, fRegistro)
            VALUES (@CodigoMozo, @Detallado, @Resumido, @Activo, @Usuario, @FechaRegistro)
            """;
        return await conn.ExecuteAsync(sql, mozo) > 0;
    }

    public async Task<bool> ActualizarAsync(Mozo mozo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE TMOZO
               SET tDetallado = @Detallado,
                   tResumido = @Resumido,
                   lActivo = @Activo,
                   tUsuario = @Usuario,
                   fRegistro = @FechaRegistro
             WHERE tCodigoMozo = @CodigoMozo
            """;
        return await conn.ExecuteAsync(sql, mozo) > 0;
    }
}
