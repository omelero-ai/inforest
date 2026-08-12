using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Infrastructure.Kitchen.Models;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Gateway de compatibilidad con SQL Legacy para cocina/KDS.
/// Legacy: USP_KDS_* + KDS_Usp_spUpdate_DPEDIDO_Orden + DPEDIDOKDS.
/// Reglas BR-010 y BR-011.
/// </summary>
internal sealed class KdsLegacyGateway : IKdsLegacyGateway
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public KdsLegacyGateway(
        IDbConnectionFactory connectionFactory,
        ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<KdsPathConfigurationRecord?> ObtenerConfiguracionAsync(CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        return await _spExecutor.QueryFirstOrDefaultAsync<KdsPathConfigurationRecord>(
            connection,
            "USP_KDS_ObtenerPath",
            cancellationToken: cancellationToken);
    }

    public async Task<IReadOnlyList<KdsDetallePedidoRecord>> ObtenerDetallePedidoAsync(
        string codigoPedido,
        int accion,
        int? kdsX = null,
        CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);

        var result = kdsX.HasValue
            ? await _spExecutor.QueryAsync<KdsDetallePedidoRecord>(
                connection,
                "USP_KDS_ObtenerDetallePedidox",
                new { tCodigoPedido = codigoPedido, accion, xkds = kdsX.Value },
                cancellationToken: cancellationToken)
            : await _spExecutor.QueryAsync<KdsDetallePedidoRecord>(
                connection,
                "USP_KDS_ObtenerDetallePedido",
                new { tCodigoPedido = codigoPedido, accion },
                cancellationToken: cancellationToken);

        return result.ToArray();
    }

    public async Task<IReadOnlyList<KdsDetalleComboRecord>> ObtenerDetalleComboAsync(
        string codigoPedido,
        string item,
        bool soloImprime,
        int accion,
        CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var result = await _spExecutor.QueryAsync<KdsDetalleComboRecord>(
            connection,
            "USP_KDS_ObtenerDetalleCombo",
            new
            {
                tCodigoPedido = codigoPedido,
                tItem = item,
                lImprime = soloImprime ? "1" : "0",
                accion
            },
            cancellationToken: cancellationToken);

        return result.ToArray();
    }

    public async Task<IReadOnlyList<KdsPropiedadProductoRecord>> ObtenerPropiedadesProductoAsync(
        string codigoPedido,
        string item,
        string? itemCombo,
        string codigoProducto,
        CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var result = await _spExecutor.QueryAsync<KdsPropiedadProductoRecord>(
            connection,
            "USP_KDS_ObtenerPropiedadesProducto",
            new
            {
                tCodigoPedido = codigoPedido,
                tItem = item,
                tItemCombo = itemCombo ?? string.Empty,
                tProducto = codigoProducto
            },
            cancellationToken: cancellationToken);

        return result.ToArray();
    }

    public async Task<IReadOnlyList<KdsAreaImpresionRecord>> ObtenerAreaImpresionKdsAsync(
        string codigoProducto,
        string? itemCombo,
        string codigoPedido,
        string item,
        CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var result = await _spExecutor.QueryAsync<KdsAreaImpresionRecord>(
            connection,
            "USP_KDS_ObtenerAreaImpresionKDS",
            new
            {
                tCodigoProducto = codigoProducto,
                tItemCombo = itemCombo ?? string.Empty,
                tCodigoPedido = codigoPedido,
                tItem = item
            },
            cancellationToken: cancellationToken);

        return result.ToArray();
    }

    public async Task SincronizarOrdenInforestAsync(
        string codigoPedido,
        int tipo,
        CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        await _spExecutor.ExecuteAsync(
            connection,
            "KDS_Usp_spUpdate_DPEDIDO_Orden",
            new
            {
                tipo,
                pedido = codigoPedido,
                tturnoKDS = codigoPedido
            },
            cancellationToken: cancellationToken);
    }

    public async Task RegistrarTiempoSalidaAsync(
        string codigoPedido,
        string item,
        DateTime fechaSalida,
        CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        await _spExecutor.ExecuteAsync(
            connection,
            "USP_KDS_GrabarTiempoSalidaDPedido",
            new
            {
                tCodigoPedido = codigoPedido,
                tItem = item,
                fSalida = fechaSalida
            },
            cancellationToken: cancellationToken);
    }
}
