using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Turno;

namespace Inforest.Infrastructure.Turno;

/// <summary>
/// Repositorio de tipo de cambio del día.
/// Legacy: TTIPOCAMBIO — spIns_TipoCambio, SELECT * FROM TTIPOCAMBIO WHERE fFecha = {fn CURDATE()}.
/// BR-TC-001.
/// </summary>
internal sealed class TipoCambioRepository : ITipoCambioRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public TipoCambioRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<TipoCambioDelDia?> ObtenerDelDiaAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // Legacy: SELECT * FROM TTIPOCAMBIO WHERE fFecha = {fn CURDATE()}
        const string sql = "SELECT nVenta, nOficial, nVenta2, nVenta3 FROM TTIPOCAMBIO WHERE fFecha = CONVERT(date, GETDATE())";
        var row = await connection.QueryFirstOrDefaultAsync<TipoCambioRow>(sql);
        if (row is null) return null;
        return new TipoCambioDelDia(
            (decimal)(row.nVenta ?? 0),
            (decimal)(row.nOficial ?? 0),
            (decimal)(row.nVenta2 ?? 0),
            (decimal)(row.nVenta3 ?? 0));
    }

    public async Task<bool> InsertarOActualizarAsync(TipoCambioDelDia tipoCambio, string usuario, CancellationToken ct = default)
    {
        // Legacy: SP spIns_TipoCambio(@nTc, @tUsuario, @nTco, @nTc2, @nTc3)
        // INSERT INTO TTIPOCAMBIO / UPDATE TTIPOCAMBIO (if exists for today). BR-TC-001.
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(connection, "spIns_TipoCambio", new
        {
            nTc     = (double)tipoCambio.Venta,
            tUSUARIO = usuario,
            nTco    = (double)tipoCambio.Oficial,
            nTc2    = (double)tipoCambio.Venta2,
            nTc3    = (double)tipoCambio.Venta3
        }, cancellationToken: ct);
        return true;
    }
}

internal sealed class TipoCambioRow
{
    public double? nVenta   { get; init; }
    public double? nOficial { get; init; }
    public double? nVenta2  { get; init; }
    public double? nVenta3  { get; init; }
}
