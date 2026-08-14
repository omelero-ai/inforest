using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Infrastructure.Ventas;

/// <summary>
/// Repositorio Dapper para notas de crédito.
/// Legacy: MNOTACREDITO. BR-NC-001..006.
/// </summary>
internal sealed class NotaCreditoRepository : INotaCreditoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public NotaCreditoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<NotaCredito?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                tNotaCredito AS CodigoNotaCredito,
                fFecha AS Fecha,
                tDocumento AS CodigoDocumento,
                ISNULL(nNeto, 0) AS Neto,
                ISNULL(nImpuesto1, 0) AS Impuesto1,
                ISNULL(nImpuesto2, 0) AS Impuesto2,
                ISNULL(nImpuesto3, 0) AS Impuesto3,
                ISNULL(nVenta, 0) AS Total,
                ISNULL(tEstadoDocumento, 'PE') AS Estado,
                tTurno AS CodigoTurno,
                tCaja AS CodigoCaja,
                tUsuario AS Usuario,
                fRegistro AS FechaRegistro,
                fDiaContable AS FechaDiaContable,
                ISNULL(tObservacion, '') AS Observacion,
                tTipoDocumento AS TipoDocumento,
                tMotivo AS Motivo,
                tUsuarioAnulado AS UsuarioAnulado,
                fRegistroAnulado AS FechaAnulado
            FROM MNOTACREDITO
            WHERE tNotaCredito = @Codigo
            """;

        var row = await connection.QueryFirstOrDefaultAsync<NotaCreditoRow>(
            new CommandDefinition(sql, new { Codigo = codigo }, cancellationToken: ct));
        return row is null ? null : Map(row);
    }

    public async Task<IReadOnlyList<NotaCredito>> ObtenerPorRangoFechaAsync(DateTime desde, DateTime hasta, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                tNotaCredito AS CodigoNotaCredito,
                fFecha AS Fecha,
                tDocumento AS CodigoDocumento,
                ISNULL(nNeto, 0) AS Neto,
                ISNULL(nImpuesto1, 0) AS Impuesto1,
                ISNULL(nImpuesto2, 0) AS Impuesto2,
                ISNULL(nImpuesto3, 0) AS Impuesto3,
                ISNULL(nVenta, 0) AS Total,
                ISNULL(tEstadoDocumento, 'PE') AS Estado,
                tTurno AS CodigoTurno,
                tCaja AS CodigoCaja,
                tUsuario AS Usuario,
                fRegistro AS FechaRegistro,
                fDiaContable AS FechaDiaContable,
                ISNULL(tObservacion, '') AS Observacion,
                tTipoDocumento AS TipoDocumento,
                tMotivo AS Motivo,
                tUsuarioAnulado AS UsuarioAnulado,
                fRegistroAnulado AS FechaAnulado
            FROM MNOTACREDITO
            WHERE fRegistro >= @Desde AND fRegistro <= @Hasta
            ORDER BY fRegistro DESC
            """;

        var rows = await connection.QueryAsync<NotaCreditoRow>(
            new CommandDefinition(sql, new { Desde = desde.Date, Hasta = hasta.Date.AddDays(1).AddSeconds(-1) }, cancellationToken: ct));
        return rows.Select(Map).ToList().AsReadOnly();
    }

    public async Task<decimal> ObtenerTotalEmitidoParaDocumentoAsync(string codigoDocumento, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT ISNULL(SUM(nVenta), 0)
            FROM MNOTACREDITO
            WHERE tDocumento = @CodigoDocumento
              AND ISNULL(tEstadoDocumento, 'PE') <> 'AN'
            """;
        return await connection.ExecuteScalarAsync<decimal>(
            new CommandDefinition(sql, new { CodigoDocumento = codigoDocumento }, cancellationToken: ct));
    }

    public async Task<bool> InsertarAsync(NotaCredito nc, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            INSERT INTO MNOTACREDITO
                (tNotaCredito, fFecha, tDocumento,
                 nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta,
                 tEstadoDocumento, tTurno, tCaja, tUsuario,
                 fRegistro, fDiaContable, tObservacion,
                 tTipoDocumento, tMotivo, lReplica)
            VALUES
                (@Codigo, @Fecha, @CodigoDocumento,
                 @Neto, @Impuesto1, @Impuesto2, @Impuesto3, @Total,
                 @Estado, @CodigoTurno, @CodigoCaja, @Usuario,
                 @FechaRegistro, @FechaDiaContable, @Observacion,
                 @TipoDocumento, @Motivo, 0)
            """;

        var rows = await connection.ExecuteAsync(
            new CommandDefinition(sql, new
            {
                Codigo = nc.CodigoNotaCredito,
                nc.Fecha,
                nc.CodigoDocumento,
                nc.Neto,
                nc.Impuesto1,
                nc.Impuesto2,
                nc.Impuesto3,
                nc.Total,
                nc.Estado,
                nc.CodigoTurno,
                nc.CodigoCaja,
                nc.Usuario,
                nc.FechaRegistro,
                nc.FechaDiaContable,
                nc.Observacion,
                nc.TipoDocumento,
                nc.Motivo
            }, cancellationToken: ct));
        return rows > 0;
    }

    public async Task<bool> AnularAsync(string codigo, string usuario, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MNOTACREDITO
            SET tEstadoDocumento = 'AN',
                tUsuarioAnulado = @Usuario,
                fRegistroAnulado = GETDATE()
            WHERE tNotaCredito = @Codigo
            """;
        var rows = await connection.ExecuteAsync(
            new CommandDefinition(sql, new { Codigo = codigo, Usuario = usuario }, cancellationToken: ct));
        return rows > 0;
    }

    private static NotaCredito Map(NotaCreditoRow r) => NotaCredito.Reconstituir(
        r.CodigoNotaCredito,
        r.Fecha,
        r.CodigoDocumento,
        r.Neto,
        r.Impuesto1,
        r.Impuesto2,
        r.Impuesto3,
        r.Total,
        r.Estado,
        r.CodigoTurno,
        r.CodigoCaja,
        r.Usuario,
        r.FechaRegistro,
        r.FechaDiaContable,
        r.Observacion,
        r.TipoDocumento,
        r.Motivo,
        r.UsuarioAnulado,
        r.FechaAnulado);

    private sealed record NotaCreditoRow(
        string CodigoNotaCredito,
        DateTime Fecha,
        string CodigoDocumento,
        decimal Neto,
        decimal Impuesto1,
        decimal Impuesto2,
        decimal Impuesto3,
        decimal Total,
        string Estado,
        string? CodigoTurno,
        string? CodigoCaja,
        string? Usuario,
        DateTime FechaRegistro,
        DateTime? FechaDiaContable,
        string Observacion,
        string? TipoDocumento,
        string? Motivo,
        string? UsuarioAnulado,
        DateTime? FechaAnulado);
}
