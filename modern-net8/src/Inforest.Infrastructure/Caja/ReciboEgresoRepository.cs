using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Repositories;

namespace Inforest.Infrastructure.Caja;

/// <summary>
/// Repositorio Dapper de recibos de egreso de caja (MEGRESO).
/// Legacy: frmReciboEgreso.frm, frmReciboEgresoDetalle.frm.
/// SP: spRep_ReciboEgreso (reporte).
/// BR-RECIBO-007, BR-RECIBO-008, BR-RECIBO-010.
/// </summary>
internal sealed class ReciboEgresoRepository : IReciboEgresoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public ReciboEgresoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    /// <inheritdoc/>
    public async Task<IReadOnlyList<ReciboEgreso>> ObtenerPorFechaAsync(
        DateTime desde,
        DateTime hasta,
        string? codigoUsuario = null,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);

        const string sqlConUsuario = """
            SELECT tRecibo, tCaja, tTurno, fFecha, tMoneda,
                   ISNULL(nTipoCambio, 1) AS nTipoCambio,
                   ISNULL(nMonto, 0) AS nMonto,
                   tDescripcion, tAutoriza,
                   ISNULL(tEstadoDocumento, '01') AS tEstadoDocumento,
                   tUsuario, fRegistro, fDiaContable, tTipoEgreso, dReferencia,
                   ISNULL(tModoEgreso, 'VARIOS') AS tModoEgreso,
                   TipoIdentificacion, Identificacion, RazonSocial,
                   NComprobante, TipoComprobante, CentroCosto, tObservacion
            FROM MEGRESO
            WHERE tUsuario = @usuario
              AND fRegistro >= @desde
              AND fRegistro <= @hasta
            ORDER BY fRegistro DESC
            """;

        const string sqlSinUsuario = """
            SELECT tRecibo, tCaja, tTurno, fFecha, tMoneda,
                   ISNULL(nTipoCambio, 1) AS nTipoCambio,
                   ISNULL(nMonto, 0) AS nMonto,
                   tDescripcion, tAutoriza,
                   ISNULL(tEstadoDocumento, '01') AS tEstadoDocumento,
                   tUsuario, fRegistro, fDiaContable, tTipoEgreso, dReferencia,
                   ISNULL(tModoEgreso, 'VARIOS') AS tModoEgreso,
                   TipoIdentificacion, Identificacion, RazonSocial,
                   NComprobante, TipoComprobante, CentroCosto, tObservacion
            FROM MEGRESO
            WHERE fRegistro >= @desde
              AND fRegistro <= @hasta
            ORDER BY fRegistro DESC
            """;

        var sql = codigoUsuario != null ? sqlConUsuario : sqlSinUsuario;
        var rows = await conn.QueryAsync<ReciboEgresoRow>(sql,
            new { usuario = codigoUsuario, desde, hasta });

        return rows.Select(MapRow).ToList();
    }

    /// <inheritdoc/>
    public async Task<ReciboEgreso?> ObtenerPorCodigoAsync(
        string codigoRecibo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tRecibo, tCaja, tTurno, fFecha, tMoneda,
                   ISNULL(nTipoCambio, 1) AS nTipoCambio,
                   ISNULL(nMonto, 0) AS nMonto,
                   tDescripcion, tAutoriza,
                   ISNULL(tEstadoDocumento, '01') AS tEstadoDocumento,
                   tUsuario, fRegistro, fDiaContable, tTipoEgreso, dReferencia,
                   ISNULL(tModoEgreso, 'VARIOS') AS tModoEgreso,
                   TipoIdentificacion, Identificacion, RazonSocial,
                   NComprobante, TipoComprobante, CentroCosto, tObservacion
            FROM MEGRESO
            WHERE tRecibo = @codigo
            """;
        var row = await conn.QueryFirstOrDefaultAsync<ReciboEgresoRow>(sql, new { codigo = codigoRecibo });
        return row is null ? null : MapRow(row);
    }

    /// <inheritdoc/>
    public async Task InsertarAsync(ReciboEgreso recibo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO MEGRESO(
                tRecibo, tCaja, tTurno, fFecha, tMoneda, nTipoCambio, nMonto,
                tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro,
                fDiaContable, tTipoEgreso, dReferencia, tModoEgreso,
                TipoIdentificacion, Identificacion, RazonSocial,
                NComprobante, TipoComprobante, CentroCosto, tObservacion)
            VALUES(
                @tRecibo, @tCaja, @tTurno, CONVERT(date, GETDATE()), @tMoneda, @nTipoCambio, @nMonto,
                @tDescripcion, @tAutoriza, @tEstadoDocumento, @tUsuario, GETDATE(),
                @fDiaContable, @tTipoEgreso, @dReferencia, @tModoEgreso,
                @TipoIdentificacion, @Identificacion, @RazonSocial,
                @NComprobante, @TipoComprobante, @CentroCosto, @tObservacion)
            """;

        await conn.ExecuteAsync(sql, new
        {
            tRecibo = recibo.CodigoRecibo,
            tCaja = recibo.CodigoCaja,
            tTurno = recibo.CodigoTurno,
            tMoneda = recibo.Moneda,
            nTipoCambio = recibo.TipoCambio,
            nMonto = recibo.Monto,
            tDescripcion = recibo.Descripcion,
            tAutoriza = recibo.Autorizacion,
            tEstadoDocumento = recibo.Estado,
            tUsuario = recibo.Usuario,
            fDiaContable = recibo.FechaDiaContable,
            tTipoEgreso = recibo.TipoEgreso,
            dReferencia = recibo.Referencia,
            tModoEgreso = recibo.ModoEgreso,
            TipoIdentificacion = recibo.TipoIdentificacion,
            Identificacion = recibo.Identificacion,
            RazonSocial = recibo.RazonSocial,
            NComprobante = recibo.NumeroComprobante,
            TipoComprobante = recibo.TipoComprobante,
            CentroCosto = recibo.CentroCosto,
            tObservacion = recibo.Observacion
        });
    }

    /// <inheritdoc/>
    public async Task ActualizarEstadoAsync(
        string codigoRecibo, string nuevoEstado, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE MEGRESO
               SET tEstadoDocumento = @estado
             WHERE tRecibo = @codigo
            """;
        await conn.ExecuteAsync(sql, new { estado = nuevoEstado, codigo = codigoRecibo });
    }

    /// <inheritdoc/>
    public async Task<string?> ObtenerMaximoCorrelativoAnioAsync(
        string anoCorto, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT MAX(tRecibo) AS Codigo
            FROM MEGRESO
            WHERE SUBSTRING(tRecibo, 1, 2) = @anoCorto
            """;
        return await conn.QueryFirstOrDefaultAsync<string?>(sql, new { anoCorto });
    }

    // ── Mapping ───────────────────────────────────────────────────────────────

    private static ReciboEgreso MapRow(ReciboEgresoRow r) =>
        ReciboEgreso.Reconstruir(
            codigo: r.tRecibo,
            caja: r.tCaja,
            turno: r.tTurno,
            fecha: r.fFecha,
            moneda: r.tMoneda ?? "01",
            tipoCambio: (decimal)r.nTipoCambio,
            monto: (decimal)r.nMonto,
            descripcion: r.tDescripcion ?? string.Empty,
            autorizacion: r.tAutoriza,
            estado: r.tEstadoDocumento ?? "01",
            usuario: r.tUsuario,
            fechaRegistro: r.fRegistro,
            fechaDiaContable: r.fDiaContable,
            tipoEgreso: r.tTipoEgreso,
            referencia: r.dReferencia,
            modoEgreso: r.tModoEgreso ?? "VARIOS",
            tipoIdentificacion: r.TipoIdentificacion,
            identificacion: r.Identificacion,
            razonSocial: r.RazonSocial,
            numeroComprobante: r.NComprobante,
            tipoComprobante: r.TipoComprobante,
            centroCosto: r.CentroCosto,
            observacion: r.tObservacion);

    private sealed class ReciboEgresoRow
    {
        public string tRecibo { get; init; } = string.Empty;
        public string? tCaja { get; init; }
        public string? tTurno { get; init; }
        public DateTime fFecha { get; init; }
        public string? tMoneda { get; init; }
        public double nTipoCambio { get; init; }
        public double nMonto { get; init; }
        public string? tDescripcion { get; init; }
        public string? tAutoriza { get; init; }
        public string? tEstadoDocumento { get; init; }
        public string? tUsuario { get; init; }
        public DateTime fRegistro { get; init; }
        public DateTime? fDiaContable { get; init; }
        public string? tTipoEgreso { get; init; }
        public string? dReferencia { get; init; }
        public string? tModoEgreso { get; init; }
        public string? TipoIdentificacion { get; init; }
        public string? Identificacion { get; init; }
        public string? RazonSocial { get; init; }
        public string? NComprobante { get; init; }
        public string? TipoComprobante { get; init; }
        public string? CentroCosto { get; init; }
        public string? tObservacion { get; init; }
    }
}
