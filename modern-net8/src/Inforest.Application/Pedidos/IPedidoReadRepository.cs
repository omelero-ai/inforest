using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Pedidos;

public interface IPedidoReadRepository
{
    Task<IReadOnlyList<Pedido>> ObtenerPorTurnoAsync(string codigoTurno, CancellationToken ct = default);
    Task<IReadOnlyList<Pedido>> ObtenerPorMesaAsync(string codigoMesa, CancellationToken ct = default);
    Task<IReadOnlyList<DetallePedido>> ObtenerDetalleAsync(string codigoPedido, CancellationToken ct = default);

    /// <summary>
    /// Obtiene items del pedido en vista enriquecida con nombre de producto, descuento% y subtotal.
    /// Legacy: vPedidoDetalle (DPEDIDO JOIN TPRODUCTO JOIN vCortesia).
    /// Regla BR-PEDIDO-005: filtra tEstadoItem = 'N' (ítems activos).
    /// </summary>
    Task<IReadOnlyList<ItemPedidoVista>> ObtenerDetalleExtendidoAsync(string codigoPedido, CancellationToken ct = default);

    /// <summary>
    /// Obtiene los ítems combo de un pedido para mostrar en el detalle.
    /// Legacy: vPedidoCombo (CPEDIDO JOIN TPRODUCTO).
    /// </summary>
    Task<IReadOnlyList<ComboPedidoVista>> ObtenerCombosAsync(string codigoPedido, CancellationToken ct = default);
}
