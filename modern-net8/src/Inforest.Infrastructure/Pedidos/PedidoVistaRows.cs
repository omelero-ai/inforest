namespace Inforest.Infrastructure.Pedidos;

// Dapper mapping record for ItemPedidoVista (ObtenerDetalleExtendidoAsync)
internal sealed class ItemPedidoVistaRow
{
    public string Item { get; init; } = string.Empty;
    public string CodigoProducto { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public decimal PrecioOficial { get; init; }
    public decimal Descuento { get; init; }
    public decimal PrecioVenta { get; init; }
    public decimal Cantidad { get; init; }
    public decimal SubTotal { get; init; }
    public string EstadoItem { get; init; } = "N";
    public string Facturado { get; init; } = "";
    public bool Imprime { get; init; }
    public bool TienePropiedad { get; init; }
    public bool TieneObservacion { get; init; }
    public bool Corte { get; init; }
    public string? Observacion { get; init; }
    public string? Documento { get; init; }
    public string? Usuario { get; init; }
    public DateTime? FechaEnvio { get; init; }
    public string? Comanda { get; init; }
}

// Dapper mapping record for ComboPedidoVista (ObtenerCombosAsync)
internal sealed class ComboPedidoVistaRow
{
    public string Item { get; init; } = string.Empty;
    public string ItemCombo { get; init; } = string.Empty;
    public string CodigoProducto { get; init; } = string.Empty;
    public string CodigoProductoCombo { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public decimal Cantidad { get; init; }
    public bool Imprime { get; init; }
    public bool TieneObservacion { get; init; }
    public bool TienePropiedad { get; init; }
    public string? Observacion { get; init; }
}
