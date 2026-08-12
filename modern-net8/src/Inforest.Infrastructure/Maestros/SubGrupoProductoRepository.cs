using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio Dapper de subgrupos de producto.
/// Legacy: TSUBGRUPOPRODUCTO / frmGrupo.frm. Database: Inforest.
/// </summary>
internal sealed class SubGrupoProductoRepository : ISubGrupoProductoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public SubGrupoProductoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<SubGrupoProducto>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoSubGrupo AS CodigoSubGrupo,
                   tCodigoGrupo AS CodigoGrupo,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro
            FROM TSUBGRUPOPRODUCTO
            ORDER BY tDetallado
            """;
        return (await conn.QueryAsync<SubGrupoProducto>(sql)).AsList();
    }

    public async Task<SubGrupoProducto?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoSubGrupo AS CodigoSubGrupo,
                   tCodigoGrupo AS CodigoGrupo,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro
            FROM TSUBGRUPOPRODUCTO
            WHERE tCodigoSubGrupo = @Codigo
            """;
        return await conn.QueryFirstOrDefaultAsync<SubGrupoProducto>(sql, new { Codigo = codigo });
    }

    public async Task<bool> InsertarAsync(SubGrupoProducto subGrupo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO TSUBGRUPOPRODUCTO (tCodigoSubGrupo, tCodigoGrupo, tDetallado, tResumido, lActivo, tUsuario, fRegistro)
            VALUES (@CodigoSubGrupo, @CodigoGrupo, @Detallado, @Resumido, @Activo, @Usuario, @FechaRegistro)
            """;
        return await conn.ExecuteAsync(sql, subGrupo) > 0;
    }

    public async Task<bool> ActualizarAsync(SubGrupoProducto subGrupo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE TSUBGRUPOPRODUCTO
               SET tCodigoGrupo = @CodigoGrupo,
                   tDetallado = @Detallado,
                   tResumido = @Resumido,
                   lActivo = @Activo,
                   tUsuario = @Usuario,
                   fRegistro = @FechaRegistro
             WHERE tCodigoSubGrupo = @CodigoSubGrupo
            """;
        return await conn.ExecuteAsync(sql, subGrupo) > 0;
    }
}
