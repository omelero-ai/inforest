using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Domain.Services;

/// <summary>
/// Servicio de dominio para filtrar productos según canal de venta.
/// Legacy: BR-014 — TPRODUCTO flags lLocal, lDelivery, lLlevar, lCanal4, lCanal5.
/// Origen: frmProducto.frm, frmProductoBoton.frm, VISIBILIDADPROPIEDADXCANAL, TVISIBILIDADTARJETACREDITOXCANAL.
/// </summary>
public sealed class ProductoVisibilidadService
{
    /// <summary>
    /// Determina si un producto es visible para un canal de venta determinado.
    /// Legacy: lLocal = visible en salón (canal 1), lDelivery (canal 2), lLlevar (canal 3), lCanal4 (canal 4), lCanal5 (canal 5).
    /// </summary>
    public bool EsVisibleParaCanal(ProductoMaestro producto, CanalVenta canal)
    {
        if (!producto.Activo) return false;
        return producto.DisponibleEnCanal(canal);
    }

    /// <summary>
    /// Filtra una colección de productos activos visibles para el canal indicado.
    /// </summary>
    public IReadOnlyList<ProductoMaestro> FiltrarPorCanal(IEnumerable<ProductoMaestro> productos, CanalVenta canal)
        => productos.Where(p => EsVisibleParaCanal(p, canal)).ToList();
}
