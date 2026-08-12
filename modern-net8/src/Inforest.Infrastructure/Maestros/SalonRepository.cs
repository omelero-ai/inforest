using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio Dapper de salones.
/// Legacy: TSALON / TTABLA(SALON) / frmMesa.frm. Database: Inforest.
/// </summary>
internal sealed class SalonRepository : ISalonRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public SalonRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<Salon>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoSalon AS CodigoSalon,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro
            FROM TSALON
            ORDER BY tDetallado
            """;
        return (await conn.QueryAsync<Salon>(sql)).AsList();
    }

    public async Task<Salon?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoSalon AS CodigoSalon,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro
            FROM TSALON
            WHERE tCodigoSalon = @Codigo
            """;
        return await conn.QueryFirstOrDefaultAsync<Salon>(sql, new { Codigo = codigo });
    }

    public async Task<bool> InsertarAsync(Salon salon, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO TSALON (tCodigoSalon, tDetallado, tResumido, lActivo, tUsuario, fRegistro)
            VALUES (@CodigoSalon, @Detallado, @Resumido, @Activo, @Usuario, @FechaRegistro)
            """;
        return await conn.ExecuteAsync(sql, salon) > 0;
    }

    public async Task<bool> ActualizarAsync(Salon salon, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE TSALON
               SET tDetallado = @Detallado,
                   tResumido = @Resumido,
                   lActivo = @Activo,
                   tUsuario = @Usuario,
                   fRegistro = @FechaRegistro
             WHERE tCodigoSalon = @CodigoSalon
            """;
        return await conn.ExecuteAsync(sql, salon) > 0;
    }
}
