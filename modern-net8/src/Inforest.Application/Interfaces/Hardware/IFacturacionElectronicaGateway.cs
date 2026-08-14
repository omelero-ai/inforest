namespace Inforest.Application.Interfaces.Hardware;

/// <summary>
/// Contrato polimórfico para la Facturación Electrónica por país.
/// Legacy: clsTrama.cls, clsxml.cls, variables lFacturacionE/tCodigoFE/CnFE en modDeclaracion.bas.
/// Regla BR-PERIPH-006. ADR-010.
/// Cada país implementa este gateway: Peru (SUNAT/OSE), Argentina (AFIP), Chile (SII), Ecuador (SRI), Bolivia (SIAT).
/// </summary>
public interface IFacturacionElectronicaGateway
{
    /// <summary>Código del país que implementa este gateway (ej. "PE", "AR", "CL", "EC", "BO").</summary>
    string CodigoPais { get; }

    /// <summary>
    /// Emite un comprobante electrónico al proveedor FE del país.
    /// Legacy: generación de XML via clsTrama/clsxml y envío a base FACTURACION MDB.
    /// </summary>
    Task<FacturacionElectronicaResult> EmitirDocumentoAsync(DocumentoFEDto documento, CancellationToken cancellationToken = default);

    /// <summary>
    /// Anula un comprobante electrónico previamente emitido.
    /// </summary>
    Task<FacturacionElectronicaResult> AnularDocumentoAsync(string idDocumento, string motivo, CancellationToken cancellationToken = default);

    /// <summary>
    /// Consulta el estado de un comprobante electrónico.
    /// </summary>
    Task<FacturacionElectronicaResult> ConsultarEstadoAsync(string idDocumento, CancellationToken cancellationToken = default);
}

/// <summary>Resultado de una operación de Facturación Electrónica.</summary>
public sealed record FacturacionElectronicaResult(
    bool EsExitoso,
    string? CodigoRespuesta,
    string? Mensaje,
    string? NumeroAutorizacion,
    string? XmlRespuesta,
    string? QrData = null)
{
    public static FacturacionElectronicaResult Exitoso(string? numeroAutorizacion, string? qrData = null, string? xmlRespuesta = null)
        => new(true, "00", "OK", numeroAutorizacion, xmlRespuesta, qrData);

    public static FacturacionElectronicaResult Fallido(string codigo, string mensaje)
        => new(false, codigo, mensaje, null, null);
}

/// <summary>
/// Datos del documento electrónico a emitir.
/// Legacy: campos de clsTrama.cls y clsxml.cls — tipo, serie, número, emisor, receptor, ítems, totales, QR.
/// </summary>
public sealed record DocumentoFEDto(
    string TipoDocumento,
    string Serie,
    string Numero,
    string CodigoPais,
    EmisorFEDto Emisor,
    ReceptorFEDto Receptor,
    IReadOnlyList<ItemFEDto> Items,
    decimal SubTotal,
    decimal Impuesto,
    decimal Total,
    string? CodigoFE = null,
    bool GenerarQR = false);

/// <summary>Datos del emisor del comprobante electrónico.</summary>
public sealed record EmisorFEDto(
    string NumeroDocumento,
    string RazonSocial,
    string Direccion,
    string TipoDocumento = "RUC");

/// <summary>Datos del receptor del comprobante electrónico.</summary>
public sealed record ReceptorFEDto(
    string NumeroDocumento,
    string RazonSocial,
    string TipoDocumento = "DNI");

/// <summary>Ítem del comprobante electrónico.</summary>
public sealed record ItemFEDto(
    string Codigo,
    string Descripcion,
    decimal Cantidad,
    decimal PrecioUnitario,
    decimal PrecioTotal,
    decimal Igv,
    string UnidadMedida = "NIU");
