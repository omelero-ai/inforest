using System.Data;
using Dapper;
using Inforest.Application.Delivery;
using Inforest.Application.Interfaces;

namespace Inforest.Infrastructure.Delivery;

/// <summary>
/// Repositorio de recargas RFID.
/// Legacy: FrmRecargarTarjeta.frm + FrmRecargarTarjetaDetalle.frm.
/// SQL Legacy: usp_Inforest_ObtieneRecargas + TMOVIMIENTOTARJETASRFID + TTARJETASRFID.
/// </summary>
internal sealed class RecargaTarjetaRepository : IRecargaTarjetaRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public RecargaTarjetaRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<RecargaTarjetaMovimiento>> ObtenerMovimientosAsync(
        DateTime desde,
        DateTime hasta,
        string tipoMovimiento,
        string? codigoRfid = null,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);

        var dp = new DynamicParameters();
        dp.Add("@FechaIni", desde.ToString("yyyy/MM/dd HH:mm"));
        dp.Add("@FechaFin", hasta.ToString("yyyy/MM/dd HH:mm"));
        dp.Add("@tipo", tipoMovimiento);
        dp.Add("@opcion", string.IsNullOrWhiteSpace(codigoRfid) ? "1" : "2");
        dp.Add("@param", string.IsNullOrWhiteSpace(codigoRfid) ? "1" : codigoRfid!.Trim());

        var rows = await conn.QueryAsync<RecargaMovimientoRow>(
            new CommandDefinition(
                "usp_Inforest_ObtieneRecargas",
                dp,
                commandType: CommandType.StoredProcedure,
                cancellationToken: ct));

        return rows.Select(MapRow).ToList();
    }

    public async Task<RecargaTarjetaResultado> RegistrarRecargaAsync(
        string codigoRfid,
        decimal montoRecarga,
        string documentoReferencia,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        using var tx = conn.BeginTransaction();
        try
        {
            const string selectSaldo = """
                SELECT CAST(ISNULL(MontoDisponible, 0) AS decimal(18,2))
                FROM TTARJETASRFID WITH (UPDLOCK, ROWLOCK)
                WHERE CodidoRFID = @CodigoRfid
                """;

            var saldoActual = await conn.QueryFirstOrDefaultAsync<decimal?>(
                new CommandDefinition(
                    selectSaldo,
                    new { CodigoRfid = codigoRfid },
                    transaction: tx,
                    cancellationToken: ct));

            if (saldoActual is null)
                throw new InvalidOperationException("No se encontró la tarjeta RFID para recargar.");

            var saldoNuevo = decimal.Round(
                saldoActual.Value + montoRecarga,
                2,
                MidpointRounding.AwayFromZero);

            const string insertMovimiento = """
                INSERT INTO TMOVIMIENTOTARJETASRFID
                    (CodidoRFID, Tipo, DocReferencia, CodReferencia, CodDispensador, Cantidad, fRegistro,
                     MontoIngreso, MontoSalida, MontoAnterior, MontoFinal)
                VALUES
                    (@CodigoRfid, 'R', @DocumentoReferencia, 1, 0, 1, GETDATE(),
                     @MontoIngreso, 0, @MontoAnterior, @MontoFinal)
                """;

            await conn.ExecuteAsync(new CommandDefinition(
                insertMovimiento,
                new
                {
                    CodigoRfid = codigoRfid,
                    DocumentoReferencia = documentoReferencia,
                    MontoIngreso = montoRecarga,
                    MontoAnterior = saldoActual.Value,
                    MontoFinal = saldoNuevo
                },
                transaction: tx,
                cancellationToken: ct));

            const string updateSaldo = """
                UPDATE TTARJETASRFID
                SET MontoDisponible = @MontoFinal,
                    FechaModificion = GETDATE()
                WHERE CodidoRFID = @CodigoRfid
                """;

            var updated = await conn.ExecuteAsync(new CommandDefinition(
                updateSaldo,
                new { CodigoRfid = codigoRfid, MontoFinal = saldoNuevo },
                transaction: tx,
                cancellationToken: ct));

            if (updated == 0)
                throw new InvalidOperationException("No se pudo actualizar el saldo de la tarjeta RFID.");

            tx.Commit();
            return new RecargaTarjetaResultado(
                codigoRfid,
                documentoReferencia,
                saldoActual.Value,
                montoRecarga,
                saldoNuevo);
        }
        catch
        {
            tx.Rollback();
            throw;
        }
    }

    private static RecargaTarjetaMovimiento MapRow(RecargaMovimientoRow row)
    {
        var fecha = DateTime.TryParse(row.FRegistro, out var parsed)
            ? parsed
            : DateTime.MinValue;

        return new RecargaTarjetaMovimiento(
            row.Id,
            row.CodigoRfid ?? string.Empty,
            row.Tipo ?? string.Empty,
            row.DocReferencia ?? string.Empty,
            fecha,
            row.MontoIngreso,
            row.MontoSalida,
            row.MontoAnterior,
            row.MontoFinal);
    }

    private sealed class RecargaMovimientoRow
    {
        public int Id { get; init; }
        public string? CodigoRfid { get; init; }
        public string? Tipo { get; init; }
        public string? DocReferencia { get; init; }
        public string? FRegistro { get; init; }
        public decimal MontoIngreso { get; init; }
        public decimal MontoSalida { get; init; }
        public decimal MontoAnterior { get; init; }
        public decimal MontoFinal { get; init; }
    }
}
