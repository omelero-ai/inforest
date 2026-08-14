namespace Inforest.Infrastructure.Pedidos;

internal sealed class DetallePedidoRecord
{
    public string CodigoPedido { get; init; } = string.Empty;
    public string Item { get; init; } = string.Empty;
    public string CodigoProducto { get; init; } = string.Empty;
    public string? CodigoGrupo { get; init; }
    public string? CodigoSubGrupo { get; init; }
    public decimal PrecioNeto { get; init; }
    public decimal PrecioVenta { get; init; }
    public decimal Cantidad { get; init; }
    public decimal Impuesto1 { get; init; }
    public decimal Impuesto2 { get; init; }
    public decimal Impuesto3 { get; init; }
    public decimal Recargo { get; init; }
    public decimal Descuento { get; init; }
    public string? Observacion { get; init; }
    public string? Area { get; init; }
    public bool Imprime { get; init; }
    public string? EstadoItem { get; init; }
}
