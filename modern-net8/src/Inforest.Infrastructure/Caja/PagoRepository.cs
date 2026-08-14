using Dapper;
using Inforest.Application.Caja;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Caja;

namespace Inforest.Infrastructure.Caja;

/// <summary>
/// Repositorio Dapper de pagos por documento.
/// Legacy operativo: DPAGODOCUMENTO.
/// </summary>
internal sealed class PagoRepository : IPagoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public PagoRepository(IDbConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory;
    }

    public async Task<bool> InsertarPagoAsync(PagoDocumento pago, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            INSERT INTO DPAGODOCUMENTO
            (
                tDocumento, tCorrelativo, tTipoPago, nMonto, tReferencia, fRegistro
            )
            VALUES
            (
                @CodigoDocumento, '001', @CodigoMedioPago, @Monto, @Referencia, @FechaRegistro
            )
            """;

        var rows = await connection.ExecuteAsync(new CommandDefinition(
            sql,
            new
            {
                pago.CodigoDocumento,
                pago.CodigoMedioPago,
                pago.Monto,
                Referencia = Truncar(pago.Referencia, 4),
                pago.FechaRegistro
            },
            cancellationToken: ct));

        return rows > 0;
    }

    public async Task<IReadOnlyList<PagoDocumento>> ObtenerPagosPorDocumentoAsync(string codDoc, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                tDocumento AS CodigoDocumento,
                tTipoPago AS CodigoMedioPago,
                nMonto AS Monto,
                CAST(0 AS decimal(18,2)) AS MontoCambio,
                tReferencia AS Referencia,
                fRegistro AS FechaRegistro
            FROM DPAGODOCUMENTO
            WHERE tDocumento = @Documento
            ORDER BY fRegistro, tCorrelativo
            """;

        var rows = await connection.QueryAsync<PagoRow>(
            new CommandDefinition(sql, new { Documento = codDoc }, cancellationToken: ct));

        return rows
            .Select(r => PagoDocumento.Reconstituir(
                r.CodigoDocumento,
                r.CodigoMedioPago,
                r.Monto,
                r.MontoCambio,
                r.Referencia,
                r.FechaRegistro))
            .ToList();
    }

    public async Task<ResumenTurnoCaja> ObtenerResumenTurnoAsync(string codigoTurno, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                @CodigoTurno AS CodigoTurno,
                CAST(ISNULL(SUM(CASE WHEN M.tEstadoDocumento = 'CO' THEN ISNULL(M.nTotal, 0) ELSE 0 END), 0) AS decimal(18,2)) AS TotalVentas,
                CAST(ISNULL(SUM(ISNULL(P.nMonto, 0)), 0) AS decimal(18,2)) AS TotalPagos,
                CAST(ISNULL(SUM(CASE WHEN M.tEstadoDocumento = 'CO' THEN ISNULL(M.nPropina, 0) ELSE 0 END), 0) AS decimal(18,2)) AS TotalPropinas,
                COUNT(DISTINCT CASE WHEN M.tEstadoDocumento = 'CO' THEN M.tDocumento END) AS DocumentosCobrados
            FROM MDOCUMENTO M
            LEFT JOIN DPAGODOCUMENTO P ON P.tDocumento = M.tDocumento
            WHERE M.tTurno = @CodigoTurno
            """;

        var resumen = await connection.QueryFirstOrDefaultAsync<ResumenTurnoCaja>(
            new CommandDefinition(sql, new { CodigoTurno = codigoTurno }, cancellationToken: ct));

        return resumen ?? new ResumenTurnoCaja(codigoTurno, 0, 0, 0, 0);
    }

    private static string? Truncar(string? value, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        var normalized = value.Trim();
        return normalized.Length <= maxLength
            ? normalized
            : normalized[..maxLength];
    }

    private sealed record PagoRow(
        string CodigoDocumento,
        string CodigoMedioPago,
        decimal Monto,
        decimal MontoCambio,
        string? Referencia,
        DateTime FechaRegistro);
}
