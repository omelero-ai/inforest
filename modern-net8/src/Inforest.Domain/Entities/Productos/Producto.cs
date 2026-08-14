using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Productos;

/// <summary>
/// Producto del sistema INFOREST — equivalente a tabla TPRODUCTO.
/// Legacy: TPRODUCTO en 1. Estructura.sql
/// Regla BR-001: 5 precios por canal de venta.
/// Regla BR-002: Hasta 15 flags de impuesto.
/// </summary>
public class Producto : Entity
{
    public string CodigoProducto { get; private set; } = string.Empty;
    public string? CodigoGrupo { get; private set; }
    public string? CodigoSubGrupo { get; private set; }
    public string Descripcion { get; private set; } = string.Empty;
    public string? DescripcionCorta { get; private set; }

    // Precios por canal de venta — Legacy: TPRODUCTO columnas nPrecioVenta*
    public decimal PrecioVenta { get; private set; }
    public decimal PrecioDelivery { get; private set; }
    public decimal PrecioLlevar { get; private set; }
    public decimal PrecioCanal4 { get; private set; }
    public decimal PrecioCanal5 { get; private set; }

    // Impuestos — Legacy: BR-002
    public decimal PorcentajeImpuesto1 { get; private set; }
    public decimal PorcentajeImpuesto2 { get; private set; }
    public decimal PorcentajeImpuesto3 { get; private set; }
    public bool AplicaImpuesto1 { get; private set; }
    public bool AplicaImpuesto2 { get; private set; }
    public bool AplicaImpuesto3 { get; private set; }

    public bool Activo { get; private set; }
    public bool TieneAlmacen { get; private set; }
    public string? Area { get; private set; }

    private Producto() { }

    /// <summary>
    /// Obtiene el precio según el canal de venta.
    /// Legacy: BR-001 — modPuntoVenta, frmVenta
    /// </summary>
    public decimal ObtenerPrecioPorCanal(int canalVenta)
    {
        return canalVenta switch
        {
            1 => PrecioVenta,
            2 => PrecioDelivery > 0 ? PrecioDelivery : PrecioVenta,
            3 => PrecioLlevar > 0 ? PrecioLlevar : PrecioVenta,
            4 => PrecioCanal4 > 0 ? PrecioCanal4 : PrecioVenta,
            5 => PrecioCanal5 > 0 ? PrecioCanal5 : PrecioVenta,
            _ => throw new DomainException($"Canal de venta inválido: {canalVenta}", "CANAL_INVALIDO")
        };
    }
}
