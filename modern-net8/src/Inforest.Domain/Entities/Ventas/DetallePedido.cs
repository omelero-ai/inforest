using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Ventas;

/// <summary>
/// Detalle de pedido — equivalente a tabla DPEDIDO.
/// Legacy: DPEDIDO en 1. Estructura.sql
/// Regla BR-001: Precio varía según canal de venta del pedido padre.
/// </summary>
public class DetallePedido : Entity
{
    public string CodigoPedido { get; private set; } = string.Empty;
    public string Item { get; private set; } = string.Empty;
    public string CodigoProducto { get; private set; } = string.Empty;
    public string? CodigoGrupo { get; private set; }
    public string? CodigoSubGrupo { get; private set; }
    public decimal PrecioNeto { get; private set; }
    public decimal PrecioVenta { get; private set; }
    public decimal Cantidad { get; private set; }
    public decimal Impuesto1 { get; private set; }
    public decimal Impuesto2 { get; private set; }
    public decimal Impuesto3 { get; private set; }
    public decimal Venta { get; private set; }
    public decimal Recargo { get; private set; }
    public decimal Descuento { get; private set; }
    public string? Observacion { get; private set; }
    public bool Imprime { get; private set; }
    public string? EstadoItem { get; private set; }
    public string? Area { get; private set; }

    private DetallePedido() { }

    public static DetallePedido Crear(
        string codigoPedido,
        string item,
        string codigoProducto,
        decimal cantidad,
        decimal precioNeto,
        decimal precioVenta,
        decimal impuesto1 = 0,
        decimal impuesto2 = 0,
        decimal impuesto3 = 0)
    {
        if (cantidad <= 0)
            throw new DomainException("La cantidad debe ser mayor a cero.", "DETALLE_CANTIDAD_INVALIDA");

        if (precioVenta < 0)
            throw new DomainException("El precio de venta no puede ser negativo.", "DETALLE_PRECIO_INVALIDO");

        var venta = cantidad * precioVenta;

        return new DetallePedido
        {
            CodigoPedido = codigoPedido,
            Item = item,
            CodigoProducto = codigoProducto,
            Cantidad = cantidad,
            PrecioNeto = precioNeto,
            PrecioVenta = precioVenta,
            Impuesto1 = impuesto1,
            Impuesto2 = impuesto2,
            Impuesto3 = impuesto3,
            Venta = venta,
            Imprime = true,
            EstadoItem = "PEN"
        };
    }
}
