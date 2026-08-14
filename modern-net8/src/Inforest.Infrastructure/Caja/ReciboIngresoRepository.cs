using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Repositories;

namespace Inforest.Infrastructure.Caja;

/// <summary>
/// Repositorio Dapper de recibos de ingreso de caja (MINGRESO / vingreso).
/// Legacy: frmReciboIngreso.frm, frmReciboIngresoDetalle.frm.
/// BR-RECIBO-001, BR-RECIBO-002, BR-RECIBO-003, BR-RECIBO-006.
/// </summary>
internal sealed class ReciboIngresoRepository : IReciboIngresoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public ReciboIngresoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    /// <inheritdoc/>
    public async Task<IReadOnlyList<ReciboIngreso>> ObtenerPorFechaAsync(
        DateTime desde,
        DateTime hasta,
        string? codigoUsuario = null,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);

        // BR-RECIBO-006: si codigoUsuario != null filtrar por usuario (lMCPV)
        const string sqlConUsuario = """
            SELECT tRecibo, fFecha, tMoneda, tTipoPago, tTarjeta, tReferencia,
                   ISNULL(nTipoCambio, 1) AS nTipoCambio,
                   ISNULL(nMonto, 0) AS nMonto,
                   tDescripcion, tAutoriza,
                   ISNULL(lAnticipo, 0) AS lAnticipo,
                   ISNULL(tEstadoDocumento, '01') AS tEstadoDocumento,
                   tTurno, tCaja, tUsuario, fRegistro, fDiaContable, tPedido
            FROM MINGRESO
            WHERE tUsuario = @usuario
              AND fRegistro >= @desde
              AND fRegistro <= @hasta
            ORDER BY fRegistro DESC
            """;

        const string sqlSinUsuario = """
            SELECT tRecibo, fFecha, tMoneda, tTipoPago, tTarjeta, tReferencia,
                   ISNULL(nTipoCambio, 1) AS nTipoCambio,
                   ISNULL(nMonto, 0) AS nMonto,
                   tDescripcion, tAutoriza,
                   ISNULL(lAnticipo, 0) AS lAnticipo,
                   ISNULL(tEstadoDocumento, '01') AS tEstadoDocumento,
                   tTurno, tCaja, tUsuario, fRegistro, fDiaContable, tPedido
            FROM MINGRESO
            WHERE fRegistro >= @desde
              AND fRegistro <= @hasta
            ORDER BY fRegistro DESC
            """;

        var sql = codigoUsuario != null ? sqlConUsuario : sqlSinUsuario;
        var rows = await conn.QueryAsync<ReciboIngresoRow>(sql,
            new { usuario = codigoUsuario, desde, hasta });

        return rows.Select(MapRow).ToList();
    }

    /// <inheritdoc/>
    public async Task<ReciboIngreso?> ObtenerPorCodigoAsync(
        string codigoRecibo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tRecibo, fFecha, tMoneda, tTipoPago, tTarjeta, tReferencia,
                   ISNULL(nTipoCambio, 1) AS nTipoCambio,
                   ISNULL(nMonto, 0) AS nMonto,
                   tDescripcion, tAutoriza,
                   ISNULL(lAnticipo, 0) AS lAnticipo,
                   ISNULL(tEstadoDocumento, '01') AS tEstadoDocumento,
                   tTurno, tCaja, tUsuario, fRegistro, fDiaContable, tPedido
            FROM MINGRESO
            WHERE tRecibo = @codigo
            """;
        var row = await conn.QueryFirstOrDefaultAsync<ReciboIngresoRow>(sql, new { codigo = codigoRecibo });
        return row is null ? null : MapRow(row);
    }

    /// <inheritdoc/>
    public async Task InsertarAsync(ReciboIngreso recibo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO MINGRESO(
                tRecibo, tCaja, tTurno, fFecha, tMoneda, nTipoCambio, nMonto,
                tTarjeta, tTipoPago, tReferencia, tDescripcion, tAutoriza,
                lAnticipo, tEstadoDocumento, tUsuario, fRegistro, fDiaContable, tPedido)
            VALUES(
                @tRecibo, @tCaja, @tTurno, @fFecha, @tMoneda, @nTipoCambio, @nMonto,
                @tTarjeta, @tTipoPago, @tReferencia, @tDescripcion, @tAutoriza,
                @lAnticipo, @tEstadoDocumento, @tUsuario, GETDATE(), @fDiaContable, @tPedido)
            """;

        await conn.ExecuteAsync(sql, new
        {
            tRecibo = recibo.CodigoRecibo,
            tCaja = recibo.CodigoCaja,
            tTurno = recibo.CodigoTurno,
            fFecha = recibo.Fecha,
            tMoneda = recibo.Moneda,
            nTipoCambio = recibo.TipoCambio,
            nMonto = recibo.Monto,
            tTarjeta = recibo.CodigoTarjeta,
            tTipoPago = recibo.TipoPago,
            tReferencia = recibo.Referencia,
            tDescripcion = recibo.Descripcion,
            tAutoriza = recibo.Autorizacion,
            lAnticipo = recibo.EsAnticipo ? 1 : 0,
            tEstadoDocumento = recibo.Estado,
            tUsuario = recibo.Usuario,
            fDiaContable = recibo.FechaDiaContable,
            tPedido = recibo.CodigoPedido
        });
    }

    /// <inheritdoc/>
    public async Task ActualizarEstadoAsync(
        string codigoRecibo, string nuevoEstado, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE MINGRESO
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
            FROM MINGRESO
            WHERE SUBSTRING(tRecibo, 1, 2) = @anoCorto
            """;
        return await conn.QueryFirstOrDefaultAsync<string?>(sql, new { anoCorto });
    }

    // ── Mapping ───────────────────────────────────────────────────────────────

    private static ReciboIngreso MapRow(ReciboIngresoRow r) =>
        ReciboIngreso.Reconstruir(
            codigo: r.tRecibo,
            fecha: r.fFecha,
            monto: (decimal)r.nMonto,
            descripcion: r.tDescripcion ?? string.Empty,
            moneda: r.tMoneda ?? "01",
            tipoPago: r.tTipoPago ?? "01",
            codigoTarjeta: r.tTarjeta,
            referencia: r.tReferencia,
            tipoCambio: (decimal)r.nTipoCambio,
            autorizacion: r.tAutoriza,
            esAnticipo: r.lAnticipo,
            estado: r.tEstadoDocumento ?? "01",
            turno: r.tTurno,
            caja: r.tCaja,
            usuario: r.tUsuario,
            fechaRegistro: r.fRegistro,
            fechaDiaContable: r.fDiaContable,
            codigoPedido: r.tPedido);

    private sealed class ReciboIngresoRow
    {
        public string tRecibo { get; init; } = string.Empty;
        public DateTime fFecha { get; init; }
        public string? tMoneda { get; init; }
        public string? tTipoPago { get; init; }
        public string? tTarjeta { get; init; }
        public string? tReferencia { get; init; }
        public double nTipoCambio { get; init; }
        public double nMonto { get; init; }
        public string? tDescripcion { get; init; }
        public string? tAutoriza { get; init; }
        public bool lAnticipo { get; init; }
        public string? tEstadoDocumento { get; init; }
        public string? tTurno { get; init; }
        public string? tCaja { get; init; }
        public string? tUsuario { get; init; }
        public DateTime fRegistro { get; init; }
        public DateTime? fDiaContable { get; init; }
        public string? tPedido { get; init; }
    }
}
