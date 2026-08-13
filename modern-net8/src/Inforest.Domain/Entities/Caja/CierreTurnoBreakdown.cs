namespace Inforest.Domain.Entities.Caja;

/// <summary>
/// Desglose de montos del cierre de turno para actualización de MTURNO.
/// Legacy: frmLiquidacionDetalle.frm — variables nEfectivoN, nEfectivoE, nChequeN, nChequeE, nPuntoN,
///         nTarjeta(1..8), nOtroMN, nOtroME, nRetiroN, nRetiroE, nIngresoN, nIngresoE, nEgresoN, nEgresoE.
/// Reglas BR-CAJA-003, BR-CAJA-004.
/// </summary>
public sealed record CierreTurnoBreakdown
{
    /// <summary>Efectivo moneda nacional (nMontoEN / nMontoFN base).</summary>
    public decimal EfectivoMN { get; init; }

    /// <summary>Efectivo moneda extranjera 1 (nMontoEE / nMontoFE).</summary>
    public decimal EfectivoME { get; init; }

    /// <summary>Efectivo moneda extranjera 2 (nMontoEE2 / nMontoFE2).</summary>
    public decimal EfectivoME2 { get; init; }

    /// <summary>Efectivo moneda extranjera 3 (nMontoEE3 / nMontoFE3).</summary>
    public decimal EfectivoME3 { get; init; }

    /// <summary>Cheque moneda nacional (nMontoCN).</summary>
    public decimal ChequesMN { get; init; }

    /// <summary>Cheque moneda extranjera (nMontoCE).</summary>
    public decimal ChequesME { get; init; }

    /// <summary>Puntos moneda nacional (nMontoPN).</summary>
    public decimal PuntosMN { get; init; }

    /// <summary>Puntos moneda extranjera (nMontoPE).</summary>
    public decimal PuntosME { get; init; }

    /// <summary>Tarjetas agrupadas por tipo (nTarjeta1..nTarjeta8). Índice 1-based, máximo 8.</summary>
    public IReadOnlyList<decimal> Tarjetas { get; init; } = Array.Empty<decimal>();

    /// <summary>Otros pagos moneda nacional (nOtroMN consolidado desde OtroTipoPago).</summary>
    public decimal OtrosMN { get; init; }

    /// <summary>Otros pagos moneda extranjera (nOtroME).</summary>
    public decimal OtrosME { get; init; }

    /// <summary>Retiro de efectivo MN (se resta de EfectivoMN en nMontoFN).</summary>
    public decimal RetiroMN { get; init; }

    /// <summary>Retiro de efectivo ME (se resta de EfectivoME en nMontoFE).</summary>
    public decimal RetiroME { get; init; }

    /// <summary>Total recibo de ingresos MN (nIngresoN).</summary>
    public decimal IngresosMN { get; init; }

    /// <summary>Total recibo de ingresos ME (nIngresoE).</summary>
    public decimal IngresosME { get; init; }

    /// <summary>Total recibo de egresos MN (nEgresoN).</summary>
    public decimal EgresosMN { get; init; }

    /// <summary>Total recibo de egresos ME (nEgresoE).</summary>
    public decimal EgresosME { get; init; }

    /// <summary>
    /// Monto final efectivo MN = EfectivoMN - RetiroMN.
    /// Equivale a nMontoFN.
    /// </summary>
    public decimal MontoFinalMN => EfectivoMN - RetiroMN;

    /// <summary>
    /// Monto final efectivo ME = EfectivoME - RetiroME.
    /// Equivale a nMontoFE.
    /// </summary>
    public decimal MontoFinalME => EfectivoME - RetiroME;

    /// <summary>Obtiene el monto de la tarjeta en posición 1-based (1..8). Retorna 0 si fuera de rango.</summary>
    public decimal ObtenerTarjeta(int indice)
        => indice >= 1 && indice <= Tarjetas.Count ? Tarjetas[indice - 1] : 0m;

    public static CierreTurnoBreakdown Vacio() => new();
}
