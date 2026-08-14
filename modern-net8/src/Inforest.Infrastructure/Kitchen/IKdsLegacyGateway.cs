using Inforest.Infrastructure.Kitchen.Models;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Gateway Legacy para SPs KDS y tablas auxiliares.
/// Legacy: USP_KDS_* + KDS_Usp_spUpdate_DPEDIDO_Orden + TAREAIMPRESORA/TPRODUCTOAREA.
/// Reglas BR-010 y BR-011.
/// </summary>
internal interface IKdsLegacyGateway
{
    Task<KdsPathConfigurationRecord?> ObtenerConfiguracionAsync(CancellationToken cancellationToken = default);

    Task<IReadOnlyList<KdsDetallePedidoRecord>> ObtenerDetallePedidoAsync(
        string codigoPedido,
        int accion,
        int? kdsX = null,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<KdsDetalleComboRecord>> ObtenerDetalleComboAsync(
        string codigoPedido,
        string item,
        bool soloImprime,
        int accion,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<KdsPropiedadProductoRecord>> ObtenerPropiedadesProductoAsync(
        string codigoPedido,
        string item,
        string? itemCombo,
        string codigoProducto,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<KdsAreaImpresionRecord>> ObtenerAreaImpresionKdsAsync(
        string codigoProducto,
        string? itemCombo,
        string codigoPedido,
        string item,
        CancellationToken cancellationToken = default);

    Task SincronizarOrdenInforestAsync(
        string codigoPedido,
        int tipo,
        CancellationToken cancellationToken = default);

    Task RegistrarTiempoSalidaAsync(
        string codigoPedido,
        string item,
        DateTime fechaSalida,
        CancellationToken cancellationToken = default);
}
