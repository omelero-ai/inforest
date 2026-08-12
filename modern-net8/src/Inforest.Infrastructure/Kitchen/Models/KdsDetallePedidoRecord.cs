namespace Inforest.Infrastructure.Kitchen.Models;

/// <summary>
/// Resultado de USP_KDS_ObtenerDetallePedido / USP_KDS_ObtenerDetallePedidox.
/// </summary>
internal sealed record KdsDetallePedidoRecord(
    string tCodigoPedido,
    string tItem,
    string? tItemCombo,
    string tCodigoProducto,
    string Producto,
    string? Combo,
    decimal nCantidad,
    bool lCombinacion);
