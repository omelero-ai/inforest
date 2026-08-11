using Inforest.Domain.Entities.Ventas;

namespace Inforest.Domain.Services;

/// <summary>
/// Servicio de dominio para cálculo de precios.
/// Legacy: Lógica distribuida en múltiples formularios VB6 (.frm) según canal de venta.
/// Regla BR-001: Canal de venta determina el precio a aplicar.
/// </summary>
public class PrecioService
{
    /// <summary>
    /// Calcula el precio neto (sin impuestos) para un producto en un canal dado.
    /// </summary>
    public decimal CalcularPrecioNeto(decimal precioVenta, decimal porcentajeImpuesto1, decimal porcentajeImpuesto2, decimal porcentajeImpuesto3)
    {
        var totalImpuesto = (porcentajeImpuesto1 + porcentajeImpuesto2 + porcentajeImpuesto3) / 100m;
        if (totalImpuesto == 0)
            return precioVenta;

        // Legacy: precio neto = precio venta / (1 + % impuesto total)
        return precioVenta / (1 + totalImpuesto);
    }

    /// <summary>
    /// Calcula el total de una línea de pedido.
    /// Legacy: frmVenta — nVenta = nCantidad * nPrecioVenta
    /// </summary>
    public decimal CalcularTotalLinea(decimal cantidad, decimal precioVenta)
        => cantidad * precioVenta;
}
