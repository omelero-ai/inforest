using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Cocina;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Lectura de configuración de cocina/KDS desde el modelo Legacy.
/// Legacy: TPARAMETRO + TAREAIMPRESORA + TPRODUCTOAREA.
/// Reglas BR-010 y BR-011.
/// </summary>
internal sealed class KitchenConfigurationProvider : IKitchenConfigurationProvider
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly IKdsLegacyGateway _legacyGateway;

    public KitchenConfigurationProvider(
        IDbConnectionFactory connectionFactory,
        IKdsLegacyGateway legacyGateway)
    {
        _connectionFactory = connectionFactory;
        _legacyGateway = legacyGateway;
    }

    public async Task<ConfiguracionProduccionCocina> ObtenerConfiguracionAsync(
        string codigoCaja,
        CancellationToken cancellationToken = default)
    {
        var legacy = await _legacyGateway.ObtenerConfiguracionAsync(cancellationToken);

        return ConfiguracionProduccionCocina.Crear(
            legacy?.lKDS ?? false,
            legacy?.lKDSInforest ?? false,
            legacy?.lKDSxArea ?? false,
            legacy?.lKDSTodos ?? false,
            legacy?.tOrderInfo,
            legacy?.tOrderInfox,
            legacy?.tBump,
            legacy?.tBumpx);
    }

    public async Task<IReadOnlyDictionary<string, string?>> ObtenerImpresorasPorAreaAsync(
        string codigoCaja,
        CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT tArea, tImpresora
            FROM TAREAIMPRESORA
            WHERE tCaja = @codigoCaja
            """;

        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var rows = await connection.QueryAsync<AreaPrinterRow>(
            new CommandDefinition(sql, new { codigoCaja }, cancellationToken: cancellationToken));

        return rows.ToDictionary(row => row.tArea, row => row.tImpresora, StringComparer.OrdinalIgnoreCase);
    }

    public async Task<IReadOnlyCollection<string>> ObtenerAreasProductoAsync(
        string codigoProducto,
        CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT tArea
            FROM TPRODUCTOAREA
            WHERE tCodigoProducto = @codigoProducto
            ORDER BY tArea
            """;

        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var rows = await connection.QueryAsync<string>(
            new CommandDefinition(sql, new { codigoProducto }, cancellationToken: cancellationToken));

        return rows.ToArray();
    }

    private sealed record AreaPrinterRow(string tArea, string? tImpresora);
}
