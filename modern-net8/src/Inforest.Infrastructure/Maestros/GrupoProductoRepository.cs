using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio Dapper de grupos de producto.
/// Legacy: TGRUPOPRODUCTO / frmGrupo.frm. Database: Inforest.
/// </summary>
internal sealed class GrupoProductoRepository : IGrupoProductoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public GrupoProductoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<GrupoProducto>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoGrupo AS CodigoGrupo,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro
            FROM TGRUPOPRODUCTO
            ORDER BY tDetallado
            """;
        return (await conn.QueryAsync<GrupoProducto>(sql)).AsList();
    }

    public async Task<GrupoProducto?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoGrupo AS CodigoGrupo,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro
            FROM TGRUPOPRODUCTO
            WHERE tCodigoGrupo = @Codigo
            """;
        return await conn.QueryFirstOrDefaultAsync<GrupoProducto>(sql, new { Codigo = codigo });
    }

    public async Task<bool> InsertarAsync(GrupoProducto grupo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO TGRUPOPRODUCTO (tCodigoGrupo, tDetallado, tResumido, lActivo, tUsuario, fRegistro)
            VALUES (@CodigoGrupo, @Detallado, @Resumido, @Activo, @Usuario, @FechaRegistro)
            """;
        return await conn.ExecuteAsync(sql, grupo) > 0;
    }

    public async Task<bool> ActualizarAsync(GrupoProducto grupo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE TGRUPOPRODUCTO
               SET tDetallado = @Detallado,
                   tResumido = @Resumido,
                   lActivo = @Activo,
                   tUsuario = @Usuario,
                   fRegistro = @FechaRegistro
             WHERE tCodigoGrupo = @CodigoGrupo
            """;
        return await conn.ExecuteAsync(sql, grupo) > 0;
    }
}
