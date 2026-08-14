using Dapper;
using Inforest.Application.Maestros;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio de cuentas corrientes.
/// Legacy: vCompania (TDELIVERY lClienteCtaCte=1) / frmCtaCte.frm / frmCuentaCobrar.frm.
/// Database: Inforest. BR-CTACTE-001/002/003.
/// </summary>
internal sealed class CuentaCorrienteRepository : ICuentaCorrienteRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public CuentaCorrienteRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<CuentaCorriente>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoDelivery                                     AS CodigoCtaCte,
                   ISNULL(tApellido, '') + ' ' + ISNULL(tNombre, '')   AS Descripcion,
                   tIdentidad                                           AS Identidad,
                   ISNULL(nLinea, 0)                                    AS LineaCredito,
                   ISNULL(nConsumo, 0)                                  AS Consumo,
                   ISNULL(nDiasCredito, 0)                              AS DiasCredito,
                   tTipoCtaCte                                          AS TipoCtaCte,
                   tCodigoCliente                                       AS CodigoCliente,
                   lActivo                                              AS Activo,
                   tUsuario                                             AS Usuario,
                   fRegistro                                            AS FechaRegistro
            FROM   TDELIVERY
            WHERE  lActivo        = 1
              AND  lClienteCtaCte = 1
            ORDER BY tApellido, tNombre
            """;
        var rows = await conn.QueryAsync<CuentaCorrienteRow>(sql);
        return rows.Select(r => CuentaCorriente.Crear(
                r.CodigoCtaCte, r.Descripcion, (decimal)r.LineaCredito, (decimal)r.Consumo,
                r.Usuario ?? string.Empty, r.Identidad, r.DiasCredito,
                r.TipoCtaCte, r.CodigoCliente, r.Activo, r.FechaRegistro))
            .ToList();
    }

    public async Task<CuentaCorriente?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoDelivery                                     AS CodigoCtaCte,
                   ISNULL(tApellido, '') + ' ' + ISNULL(tNombre, '')   AS Descripcion,
                   tIdentidad                                           AS Identidad,
                   ISNULL(nLinea, 0)                                    AS LineaCredito,
                   ISNULL(nConsumo, 0)                                  AS Consumo,
                   ISNULL(nDiasCredito, 0)                              AS DiasCredito,
                   tTipoCtaCte                                          AS TipoCtaCte,
                   tCodigoCliente                                       AS CodigoCliente,
                   lActivo                                              AS Activo,
                   tUsuario                                             AS Usuario,
                   fRegistro                                            AS FechaRegistro
            FROM   TDELIVERY
            WHERE  tCodigoDelivery = @Codigo
            """;
        var r = await conn.QueryFirstOrDefaultAsync<CuentaCorrienteRow>(sql, new { Codigo = codigo });
        return r is null ? null : CuentaCorriente.Crear(
            r.CodigoCtaCte, r.Descripcion, (decimal)r.LineaCredito, (decimal)r.Consumo,
            r.Usuario ?? string.Empty, r.Identidad, r.DiasCredito,
            r.TipoCtaCte, r.CodigoCliente, r.Activo, r.FechaRegistro);
    }

    public async Task<IReadOnlyList<DocumentoPendienteCobro>> ObtenerDocumentosPendientesCobroAsync(
        DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // BR-CTACTE-002: estado '03' = documento en cuenta corriente pendiente de cobro.
        const string sql = """
            SELECT tDocumento          AS Documento,
                   tCaja               AS Caja,
                   tCodigoCliente      AS CodigoCliente,
                   Cliente             AS NombreCliente,
                   tClientePago        AS CodigoClientePago,
                   ClientePago         AS NombreClientePago,
                   nVenta              AS Venta,
                   fFecha              AS Fecha,
                   tTurno              AS Turno,
                   tUsuario            AS Usuario,
                   Vencimiento         AS Vencimiento
            FROM   vDocumentoGrilla
            WHERE  fFecha          >= @FechaInicio
              AND  fFecha          <= @FechaFin
              AND  tEstadoDocumento = '03'
            ORDER BY tDocumento
            """;
        var rows = await conn.QueryAsync<DocumentoPendienteRow>(sql, new
        {
            FechaInicio = fechaInicio.Date,
            FechaFin    = fechaFin.Date.AddDays(1).AddSeconds(-1)
        });
        return rows.Select(r => new DocumentoPendienteCobro(
                r.Documento, r.Caja, r.CodigoCliente, r.NombreCliente,
                r.CodigoClientePago, r.NombreClientePago, (decimal)r.Venta,
                r.Fecha, r.Turno, r.Usuario, r.Vencimiento))
            .ToList();
    }

    // ── internal row projections ────────────────────────────────────────────
    private sealed class CuentaCorrienteRow
    {
        public string CodigoCtaCte { get; init; } = string.Empty;
        public string Descripcion { get; init; } = string.Empty;
        public string? Identidad { get; init; }
        public double LineaCredito { get; init; }
        public double Consumo { get; init; }
        public int DiasCredito { get; init; }
        public string? TipoCtaCte { get; init; }
        public string? CodigoCliente { get; init; }
        public bool Activo { get; init; }
        public string? Usuario { get; init; }
        public DateTime FechaRegistro { get; init; }
    }

    private sealed class DocumentoPendienteRow
    {
        public string Documento { get; init; } = string.Empty;
        public string Caja { get; init; } = string.Empty;
        public string? CodigoCliente { get; init; }
        public string? NombreCliente { get; init; }
        public string? CodigoClientePago { get; init; }
        public string? NombreClientePago { get; init; }
        public double Venta { get; init; }
        public DateTime Fecha { get; init; }
        public string Turno { get; init; } = string.Empty;
        public string? Usuario { get; init; }
        public string? Vencimiento { get; init; }
    }
}
