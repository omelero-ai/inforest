using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Almacen;

/// <summary>
/// Línea de detalle de un requerimiento de almacén.
/// Legacy: vRequerimiento (CodArt, Articulo, nAtendido, tCodigoProducto).
/// frmImportacionRequerimientoDetalle.frm.
/// BR-IMPORT-003: El enlace tEnlace con INFOREST es obligatorio para importar.
/// </summary>
public sealed class DetalleRequerimientoAlmacen
{
    /// <summary>Código de artículo en ALMACEN.</summary>
    public string CodArt { get; }

    /// <summary>Descripción del artículo.</summary>
    public string Articulo { get; }

    /// <summary>Cantidad aprobada/atendida en el requerimiento.</summary>
    public double Cantidad { get; }

    /// <summary>
    /// Código del producto equivalente en INFOREST (vProducto.tCodigoProducto).
    /// Legacy: TPRODUCTO.tEnlace = CodArt de ALMACEN.
    /// Null o vacío si no existe enlace.
    /// </summary>
    public string? EnlaceInforRest { get; }

    /// <summary>
    /// BR-IMPORT-003: El producto debe tener un enlace válido con INFOREST.
    /// Si no hay enlace, el pedido debe cancelarse.
    /// </summary>
    public bool TieneEnlace => !string.IsNullOrWhiteSpace(EnlaceInforRest);

    private DetalleRequerimientoAlmacen(
        string codArt,
        string articulo,
        double cantidad,
        string? enlaceInforRest)
    {
        CodArt = codArt;
        Articulo = articulo;
        Cantidad = cantidad;
        EnlaceInforRest = enlaceInforRest;
    }

    /// <summary>
    /// Crea un detalle de requerimiento de almacén.
    /// </summary>
    public static DetalleRequerimientoAlmacen Crear(
        string codArt,
        string articulo,
        double cantidad,
        string? enlaceInforRest = null)
    {
        if (string.IsNullOrWhiteSpace(codArt))
            throw new DomainException("El código de artículo no puede estar vacío.", "DETALLE_REQ_CODART_REQUERIDO");
        if (cantidad <= 0)
            throw new DomainException("La cantidad del requerimiento debe ser mayor a cero.", "DETALLE_REQ_CANTIDAD_INVALIDA");

        return new DetalleRequerimientoAlmacen(
            codArt.Trim(),
            articulo?.Trim() ?? string.Empty,
            cantidad,
            string.IsNullOrWhiteSpace(enlaceInforRest) ? null : enlaceInforRest.Trim());
    }
}
