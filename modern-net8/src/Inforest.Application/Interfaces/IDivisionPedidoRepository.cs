using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Contrato de persistencia para la División de Pedidos.
/// Legacy: operaciones directas sobre DPEDIDO, CPEDIDO, MPEDIDO en frmDivision.frm.
/// </summary>
public interface IDivisionPedidoRepository
{
    /// <summary>
    /// Carga los items del pedido origen para iniciar la sesión de división.
    /// Equivale a INSERT inicial en sDetalle + consulta para grdDetalle1.
    /// </summary>
    Task<IEnumerable<ItemDivisionRecord>> CargarItemsPedidoAsync(
        string codigoPedido,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Obtiene el estado del pedido origen para validación BR-DIV-001.
    /// </summary>
    Task<string?> ObtenerEstadoPedidoAsync(
        string codigoPedido,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Obtiene el monto máximo configurado en el pedido (nMontoMaximo).
    /// BR-DIV-002.
    /// </summary>
    Task<decimal> ObtenerMontoMaximoPedidoAsync(
        string codigoPedido,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Obtiene la suma de nVenta del DPEDIDO real (para validar montoMaximo).
    /// BR-DIV-002.
    /// </summary>
    Task<decimal> ObtenerVentaActualPedidoAsync(
        string codigoPedido,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Obtiene el tipo de pedido del MPEDIDO origen.
    /// </summary>
    Task<string> ObtenerTipoPedidoAsync(
        string codigoPedido,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Genera el siguiente código correlativo de MPEDIDO para el año actual.
    /// Legacy: MAX(tCodigoPedido) WHERE substring(tCodigoPedido,1,2)=YY.
    /// </summary>
    Task<string> GenerarSiguienteCorrelativoPedidoAsync(
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Persiste los resultados de la división en una transacción:
    /// - DELETE + INSERT DPEDIDO para el pedido origen
    /// - Llama spUpdate_DPEDIDO_Ina para el origen
    /// - DELETE + INSERT CPEDIDO para el origen
    /// - Para cada pedido destino con items: INSERT MPEDIDO + INSERT DPEDIDO + SP
    /// </summary>
    Task ConfirmarDivisionAsync(
        ConfirmarDivisionParams parametros,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifica si el pedido tiene propiedades en TPRODUCTOPROPIEDAD para un item.
    /// </summary>
    Task<bool> TieneItemPropiedadAsync(
        string codigoPedido,
        string item,
        CancellationToken cancellationToken = default);
}

/// <summary>Fila proyectada de DPEDIDO para la sesión de división.</summary>
public sealed record ItemDivisionRecord(
    string CodigoPedido,
    string Item,
    string TipoPedido,
    string CodigoProducto,
    string NombreProducto,
    string CodigoGrupo,
    string CodigoSubGrupo,
    string Moneda,
    double PrecioNeto,
    double PrecioImpuesto1,
    double PrecioImpuesto2,
    double PrecioImpuesto3,
    double PrecioVenta,
    double Recargo,
    double Descuento,
    double PrecioOficial,
    double Cantidad,
    double Impuesto1,
    double Impuesto2,
    double Impuesto3,
    double Venta,
    string? Observacion,
    string? Cortesia,
    bool Imprime,
    string? EstadoItem,
    string? Area);

/// <summary>Parámetros para confirmar la división en la base de datos.</summary>
public sealed record ConfirmarDivisionParams(
    string CodigoPedidoOrigen,
    string TipoPedido,
    string FechaDiaContable,
    IReadOnlyList<ItemDivision> ItemsOrigen,
    IReadOnlyList<(PedidoDestino Pedido, IReadOnlyList<ItemDivision> Items)> PedidosDestino,
    bool TieneDisgregados);
