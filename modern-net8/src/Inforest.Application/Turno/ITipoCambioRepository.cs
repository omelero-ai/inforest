namespace Inforest.Application.Turno;

/// <summary>
/// Repositorio de tipo de cambio del día.
/// Legacy: tabla TTIPOCAMBIO, SP spIns_TipoCambio.
/// </summary>
public interface ITipoCambioRepository
{
    /// <summary>
    /// Obtiene el tipo de cambio del día actual.
    /// Legacy: SELECT * FROM TTIPOCAMBIO WHERE fFecha = {fn CURDATE()}
    /// </summary>
    Task<TipoCambioDelDia?> ObtenerDelDiaAsync(CancellationToken ct = default);

    /// <summary>
    /// Inserta o actualiza el tipo de cambio del día.
    /// Legacy: spIns_TipoCambio(@nTc, @tUsuario, @nTco, @nTc2, @nTc3)
    /// </summary>
    Task<bool> InsertarOActualizarAsync(TipoCambioDelDia tipoCambio, string usuario, CancellationToken ct = default);
}

/// <summary>
/// DTO de tipo de cambio del día.
/// Legacy: TTIPOCAMBIO — fFecha, nVenta, nCompra, nOficial, nVenta2, nVenta3.
/// </summary>
public sealed record TipoCambioDelDia(
    decimal Venta,      // nVenta  — tipo de cambio venta (nTC)
    decimal Oficial,    // nOficial — tipo de cambio oficial (nTCO)
    decimal Venta2,     // nVenta2 — TC adicional para Bolivia (nTC2)
    decimal Venta3      // nVenta3 — TC adicional para Bolivia (nTC3)
)
{
    public static TipoCambioDelDia Vacio() => new(0m, 0m, 0m, 0m);

    /// <summary>
    /// True cuando todos los tipos de cambio requeridos para la moneda extranjera están ingresados.
    /// Legacy: frmInicio.frm — validación antes de aperturar.
    /// </summary>
    public bool EsCompleto(bool tieneMonedaExtranjera, string codigoPais)
    {
        if (!tieneMonedaExtranjera) return true;
        if (codigoPais == "003")
            return Venta > 0 && Venta2 > 0 && Venta3 > 0;
        return Venta > 0;
    }
}
