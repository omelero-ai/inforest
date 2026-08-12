using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Ventas;

/// <summary>
/// Detalle del documento de venta.
/// Legacy: DDOCUMENTO.
/// </summary>
public class DetalleDocumento
{
    public string CodigoDocumento { get; private set; } = string.Empty;
    public string Item { get; private set; } = string.Empty;
    public string CodigoPedido { get; private set; } = string.Empty;
    public string CodigoProducto { get; private set; } = string.Empty;
    public decimal Cantidad { get; private set; }
    public decimal PrecioNeto { get; private set; }
    public decimal PrecioVenta { get; private set; }
    public decimal Impuesto1 { get; private set; }
    public decimal Impuesto2 { get; private set; }
    public decimal Impuesto3 { get; private set; }
    public decimal Venta { get; private set; }
    public decimal Descuento { get; private set; }
    public string? Area { get; private set; }
    public string? UnidadNegocio { get; private set; }
    public DateTime FechaRegistro { get; private set; }

    private DetalleDocumento()
    {
    }

    public static DetalleDocumento Crear(
        string codigoDocumento,
        string item,
        string codigoPedido,
        string codigoProducto,
        decimal cantidad,
        decimal precioNeto,
        decimal precioVenta,
        decimal impuesto1 = 0,
        decimal impuesto2 = 0,
        decimal impuesto3 = 0,
        decimal descuento = 0,
        string? area = null,
        string? unidadNegocio = null)
    {
        if (string.IsNullOrWhiteSpace(codigoDocumento))
            throw new DomainException("El código de documento es obligatorio.", "DETALLE_DOCUMENTO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(item))
            throw new DomainException("El item del detalle es obligatorio.", "DETALLE_ITEM_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoProducto))
            throw new DomainException("El producto del detalle es obligatorio.", "DETALLE_PRODUCTO_REQUERIDO");

        if (cantidad <= 0)
            throw new DomainException("La cantidad del detalle debe ser mayor a cero.", "DETALLE_DOCUMENTO_CANTIDAD_INVALIDA");

        if (precioNeto < 0 || precioVenta < 0 || impuesto1 < 0 || impuesto2 < 0 || impuesto3 < 0 || descuento < 0)
            throw new DomainException("Los importes del detalle no pueden ser negativos.", "DETALLE_DOCUMENTO_MONTO_INVALIDO");

        return new DetalleDocumento
        {
            CodigoDocumento = codigoDocumento.Trim(),
            Item = item.Trim(),
            CodigoPedido = codigoPedido.Trim(),
            CodigoProducto = codigoProducto.Trim(),
            Cantidad = cantidad,
            PrecioNeto = precioNeto,
            PrecioVenta = precioVenta,
            Impuesto1 = impuesto1,
            Impuesto2 = impuesto2,
            Impuesto3 = impuesto3,
            Venta = cantidad * precioVenta,
            Descuento = descuento,
            Area = area,
            UnidadNegocio = unidadNegocio,
            FechaRegistro = DateTime.Now
        };
    }

    public static DetalleDocumento Reconstituir(
        string codigoDocumento,
        string item,
        string codigoPedido,
        string codigoProducto,
        decimal cantidad,
        decimal precioNeto,
        decimal precioVenta,
        decimal impuesto1,
        decimal impuesto2,
        decimal impuesto3,
        decimal venta,
        decimal descuento,
        string? area,
        string? unidadNegocio,
        DateTime fechaRegistro)
    {
        return new DetalleDocumento
        {
            CodigoDocumento = codigoDocumento,
            Item = item,
            CodigoPedido = codigoPedido,
            CodigoProducto = codigoProducto,
            Cantidad = cantidad,
            PrecioNeto = precioNeto,
            PrecioVenta = precioVenta,
            Impuesto1 = impuesto1,
            Impuesto2 = impuesto2,
            Impuesto3 = impuesto3,
            Venta = venta,
            Descuento = descuento,
            Area = area,
            UnidadNegocio = unidadNegocio,
            FechaRegistro = fechaRegistro
        };
    }
}
