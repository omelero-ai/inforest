using Inforest.Domain.Entities.Configuracion;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Domain.Services;

/// <summary>
/// Política de impuestos por producto.
/// Legacy: BR-002 — TPRODUCTO flags lImpuesto1..15 + TPARAMETRO Impuesto1/2/3.
/// Origen: modDeclaracion.bas (nPorcentaje1/2/3), TPRODUCTO (lImpuesto1..15).
/// </summary>
public sealed class TaxPolicy
{
    /// <summary>
    /// Calcula el monto de impuesto 1 para un producto dado su precio de venta.
    /// Legacy: nImpuesto1 = nPrecioVenta * (nPorcentaje1 / 100) si lImpuesto1 == true.
    /// </summary>
    public static decimal CalcularImpuesto1(ProductoMaestro producto, ConfiguracionSistema configuracion, decimal precioVenta)
    {
        if (!producto.Impuesto1) return 0m;
        return Round(precioVenta * (decimal)configuracion.Impuesto1 / 100m);
    }

    /// <summary>
    /// Calcula el monto de impuesto 2 para un producto dado su precio de venta.
    /// </summary>
    public static decimal CalcularImpuesto2(ProductoMaestro producto, ConfiguracionSistema configuracion, decimal precioVenta)
    {
        if (!producto.Impuesto2) return 0m;
        return Round(precioVenta * (decimal)configuracion.Impuesto2 / 100m);
    }

    /// <summary>
    /// Calcula el monto de impuesto 3 para un producto dado su precio de venta.
    /// </summary>
    public static decimal CalcularImpuesto3(ProductoMaestro producto, ConfiguracionSistema configuracion, decimal precioVenta)
    {
        if (!producto.Impuesto3) return 0m;
        return Round(precioVenta * (decimal)configuracion.Impuesto3 / 100m);
    }

    /// <summary>
    /// Calcula el precio neto (sin impuestos) de un precio de venta con impuestos incluidos.
    /// Legacy: nNeto = nPrecioVenta / (1 + %Impuesto/100).
    /// </summary>
    public static decimal CalcularPrecioNeto(decimal precioVenta, decimal porcentajeImpuesto1, decimal porcentajeImpuesto2, decimal porcentajeImpuesto3)
    {
        var totalImpuesto = (porcentajeImpuesto1 + porcentajeImpuesto2 + porcentajeImpuesto3) / 100m;
        if (totalImpuesto == 0) return precioVenta;
        return Round(precioVenta / (1m + totalImpuesto));
    }

    /// <summary>
    /// Devuelve el porcentaje total de impuestos aplicables a un producto.
    /// </summary>
    public static decimal ObtenerPorcentajeTotalImpuestos(ProductoMaestro producto, ConfiguracionSistema configuracion)
    {
        var pct = 0m;
        if (producto.Impuesto1) pct += (decimal)configuracion.Impuesto1;
        if (producto.Impuesto2) pct += (decimal)configuracion.Impuesto2;
        if (producto.Impuesto3) pct += (decimal)configuracion.Impuesto3;
        return pct;
    }

    private static decimal Round(decimal value) => Math.Round(value, 4, MidpointRounding.AwayFromZero);
}
