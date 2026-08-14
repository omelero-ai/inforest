using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio Dapper de mesas.
/// Legacy: TMESA / frmMesa.frm / frmMesas.frm. Database: Inforest.
/// </summary>
internal sealed class MesaRepository : IMesaRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public MesaRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<Mesa>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoMesa AS CodigoMesa,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   tSalon AS CodigoSalon,
                   lFumador AS Fumador,
                   tX AS PosicionX,
                   tY AS PosicionY,
                   nPersona AS NumeroPersonas,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro,
                   CASE tEstadoMesa
                       WHEN '02' THEN 1
                       WHEN '03' THEN 2
                       WHEN '04' THEN 3
                       ELSE 0
                   END AS Estado
            FROM TMESA
            ORDER BY tSalon, tCodigoMesa
            """;
        return (await conn.QueryAsync<Mesa>(sql)).AsList();
    }

    public async Task<Mesa?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoMesa AS CodigoMesa,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   tSalon AS CodigoSalon,
                   lFumador AS Fumador,
                   tX AS PosicionX,
                   tY AS PosicionY,
                   nPersona AS NumeroPersonas,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro,
                   CASE tEstadoMesa
                       WHEN '02' THEN 1
                       WHEN '03' THEN 2
                       WHEN '04' THEN 3
                       ELSE 0
                   END AS Estado
            FROM TMESA
            WHERE tCodigoMesa = @Codigo
            """;
        return await conn.QueryFirstOrDefaultAsync<Mesa>(sql, new { Codigo = codigo });
    }

    public async Task<bool> InsertarAsync(Mesa mesa, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO TMESA (tCodigoMesa, tDetallado, tResumido, tSalon, lFumador, tX, tY, nPersona, lActivo, tUsuario, fRegistro, tEstadoMesa)
            VALUES (@CodigoMesa, @Detallado, @Resumido, @CodigoSalon, @Fumador, @PosicionX, @PosicionY, @NumeroPersonas, @Activo, @Usuario, @FechaRegistro, @EstadoMesa)
            """;
        return await conn.ExecuteAsync(sql, new
        {
            mesa.CodigoMesa,
            mesa.Detallado,
            mesa.Resumido,
            CodigoSalon = mesa.CodigoSalon,
            mesa.Fumador,
            PosicionX = mesa.PosicionX,
            PosicionY = mesa.PosicionY,
            NumeroPersonas = mesa.NumeroPersonas,
            mesa.Activo,
            mesa.Usuario,
            FechaRegistro = mesa.FechaRegistro,
            EstadoMesa = ToLegacyEstado(mesa.Estado)
        }) > 0;
    }

    public async Task<bool> ActualizarAsync(Mesa mesa, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE TMESA
               SET tDetallado = @Detallado,
                   tResumido = @Resumido,
                   tSalon = @CodigoSalon,
                   lFumador = @Fumador,
                   tX = @PosicionX,
                   tY = @PosicionY,
                   nPersona = @NumeroPersonas,
                   lActivo = @Activo,
                   tUsuario = @Usuario,
                   fRegistro = @FechaRegistro,
                   tEstadoMesa = @EstadoMesa
             WHERE tCodigoMesa = @CodigoMesa
            """;
        return await conn.ExecuteAsync(sql, new
        {
            mesa.CodigoMesa,
            mesa.Detallado,
            mesa.Resumido,
            CodigoSalon = mesa.CodigoSalon,
            mesa.Fumador,
            PosicionX = mesa.PosicionX,
            PosicionY = mesa.PosicionY,
            NumeroPersonas = mesa.NumeroPersonas,
            mesa.Activo,
            mesa.Usuario,
            FechaRegistro = mesa.FechaRegistro,
            EstadoMesa = ToLegacyEstado(mesa.Estado)
        }) > 0;
    }

    private static string ToLegacyEstado(EstadoMesa estado)
        => estado switch
        {
            EstadoMesa.Ocupada => "02",
            EstadoMesa.Reservada => "03",
            EstadoMesa.EnCuenta => "04",
            _ => "01"
        };
}
