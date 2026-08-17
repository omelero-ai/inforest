namespace Inforest.Application.Pedidos;

/// <summary>
/// Vista enriquecida de un ítem de pedido para presentación en formularios.
/// Legacy: vPedidoDetalle view — DPEDIDO JOIN TPRODUCTO JOIN vCortesia.
/// Regla BR-PEDIDO-005: El detalle muestra solo ítems activos (tEstadoItem='N').
/// </summary>
public sealed record ItemPedidoVista(
    string Item,
    string CodigoProducto,
    string Producto,
    decimal PrecioOficial,
    decimal Descuento,
    decimal PrecioVenta,
    decimal Cantidad,
    decimal SubTotal,
    string EstadoItem,
    bool Facturado,
    bool Imprime,
    bool TienePropiedad,
    bool TieneObservacion,
    bool Corte,
    string? Observacion,
    string? Documento,
    string? Usuario,
    DateTime? FechaEnvio,
    string? Comanda);

/// <summary>
/// Vista de un ítem combo asociado a un pedido.
/// Legacy: vPedidoCombo view — CPEDIDO JOIN TPRODUCTO.
/// </summary>
public sealed record ComboPedidoVista(
    string Item,
    string ItemCombo,
    string CodigoProducto,
    string CodigoProductoCombo,
    string Producto,
    decimal Cantidad,
    bool Imprime,
    bool TieneObservacion,
    bool TienePropiedad,
    string? Observacion);

/// <summary>
/// Resultado agregado del handler ObtenerDetallePedidoExtendidoHandler.
/// </summary>
public sealed record DetallePedidoExtendidoResultado(
    string CodigoPedido,
    int NumeroAdultos,
    int NumeroNinos,
    DateTime FechaRegistro,
    string? Observacion,
    string? ClienteDelivery,
    decimal Total,
    IReadOnlyList<ItemPedidoVista> Items,
    IReadOnlyList<ComboPedidoVista> Combos);
