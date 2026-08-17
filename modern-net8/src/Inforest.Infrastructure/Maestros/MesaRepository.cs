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
                       WHEN '05' THEN 4
                       WHEN '06' THEN 5
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
                       WHEN '05' THEN 4
                       WHEN '06' THEN 5
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
            EstadoMesa.Sucia => "04",
            EstadoMesa.Bloqueada => "05",
            EstadoMesa.FueraDeServicio => "06",
            EstadoMesa.EnCuenta => "02",  // en cuenta: sigue ocupada en legacy
            _ => "01"
        };

    /// <summary>
    /// Cambia el estado de una mesa directamente.
    /// Legacy: UPDATE TMESA SET tEstadoMesa='XX' WHERE tCodigoMesa='...' (frmMesaConsulta.frm).
    /// BR-MESACONSULTA-001.
    /// </summary>
    public async Task<bool> CambiarEstadoAsync(string codigoMesa, EstadoMesa nuevoEstado, CancellationToken ct = default)
    {
        using var cn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = "UPDATE TMESA SET tEstadoMesa = @Estado WHERE tCodigoMesa = @CodigoMesa";
        var rows = await cn.ExecuteAsync(sql, new
        {
            Estado    = ToLegacyEstado(nuevoEstado),
            CodigoMesa = codigoMesa
        });
        return rows > 0;
    }
}
