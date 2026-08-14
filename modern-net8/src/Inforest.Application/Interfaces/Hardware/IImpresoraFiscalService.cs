namespace Inforest.Application.Interfaces.Hardware;

/// <summary>
/// Abstracción para la impresora fiscal Epson (Argentina).
/// Legacy: modImpresoraFiscal.bas — IFEpson.ocx (SetPaperSize, SetPreference, SetZone, ítems de factura).
/// Regla BR-PERIPH-005. ADR-009.
/// </summary>
public interface IImpresoraFiscalService
{
    /// <summary>
    /// Configura el formulario fiscal (zonas, tamaño de papel, preferencias).
    /// Legacy: SeteoFactura(PrinterOCX) — SetPaperSize, SetPreference, SetZone.
    /// </summary>
    Task<bool> ConfigurarAsync(ConfiguracionImpresoraFiscal configuracion, CancellationToken cancellationToken = default);

    /// <summary>
    /// Emite un documento fiscal (factura, ticket fiscal).
    /// Legacy: emisión a través de PrinterFiscal OCX con ítems de venta.
    /// </summary>
    Task<bool> EmitirFacturaAsync(FiscalDocumentoDto datosFactura, CancellationToken cancellationToken = default);

    /// <summary>
    /// Cierra la jornada fiscal del día.
    /// Legacy: cierre de caja fiscal diario (requerido por regulación AFIP Argentina).
    /// </summary>
    Task<bool> CerrarJornadaAsync(CancellationToken cancellationToken = default);
}

/// <summary>Configuración de la impresora fiscal.</summary>
public sealed record ConfiguracionImpresoraFiscal(
    string Puerto,
    int ColumnasPapel = 90,
    int FilasPapel = 72);

/// <summary>
/// Datos del documento fiscal para emisión.
/// Legacy: zonas 10 (letra), 24 (razón social), 34 (nro/fecha), 36 (CUIT), 42/52 (comprador), 60-65 (ítems).
/// </summary>
public sealed record FiscalDocumentoDto(
    string LetraComprobante,
    string RazonSocial,
    string RazonComercial,
    string Cuit,
    string Direccion,
    string NumeroFactura,
    DateOnly FechaEmision,
    string? CuitComprador,
    string? RazonSocialComprador,
    IReadOnlyList<FiscalItemDto> Items,
    decimal Total);

/// <summary>Ítem de un documento fiscal.</summary>
public sealed record FiscalItemDto(
    decimal Cantidad,
    string Descripcion,
    decimal PrecioUnitario,
    decimal TasaIva,
    decimal TasaAjuste = 0);
