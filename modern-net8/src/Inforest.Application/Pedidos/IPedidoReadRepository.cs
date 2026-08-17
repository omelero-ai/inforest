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

    /// <summary>
    /// Obtiene pedidos activos (tEstadoPedido='01') sin mesa asignada de una caja.
    /// Legacy: frmMesas.frm — MPEDIDO WHERE tEstadoPedido='01' AND tTipoPedido &lt;&gt; '04' AND LEN(RTRIM(tMesa))=0 AND tCaja=@caja.
    /// Regla BR-MESAS-005.
    /// </summary>
    Task<IReadOnlyList<PedidoSinMesaVista>> ObtenerActivosSinMesaAsync(string caja, CancellationToken ct = default);

    /// <summary>
    /// Obtiene el correlativo de pedidos por rango de fechas.
    /// Legacy: frmPedidoCorrelativo.frm — vPedidoCorrelativo WHERE fFecha BETWEEN @desde AND @hasta ORDER BY tCodigoPedido.
    /// Regla BR-CORRPEDIDO-001.
    /// </summary>
    Task<IReadOnlyList<PedidoCorrelativoVista>> ObtenerCorrelativoAsync(DateOnly desde, DateOnly hasta, CancellationToken ct = default);

    /// <summary>
    /// Obtiene los documentos agrupados de un pedido (documentos refacturados/anulados).
    /// Legacy: frmPedidoAnterior.frm — vDocumentoAgrupado WHERE tCodigoPedido = @codigoPedido.
    /// </summary>
    Task<IReadOnlyList<DocumentoAgrupadoVista>> ObtenerDocumentosAgrupadosPedidoAsync(string codigoPedido, CancellationToken ct = default);
}
