using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Country;

/// <summary>
/// Gateway de Facturación Electrónica para Argentina (AFIP WS-FE).
/// Legacy: modImpresoraFiscal.bas — IFEpson.ocx Argentina, CUIT, AFIP.
/// Regla BR-PERIPH-005, BR-PERIPH-006. ADR-010.
/// </summary>
internal sealed class ArgentinaFEGateway : IFacturacionElectronicaGateway
{
    public string CodigoPais => "AR";
    private readonly ILogger<ArgentinaFEGateway> _logger;
    private readonly ArgentinaFEOptions _options;

    public ArgentinaFEGateway(ILogger<ArgentinaFEGateway> logger, Microsoft.Extensions.Options.IOptions<ArgentinaFEOptions> options)
    {
        _logger = logger;
        _options = options.Value;
    }

    public async Task<FacturacionElectronicaResult> EmitirDocumentoAsync(DocumentoFEDto documento, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("FE Argentina (AFIP): emitiendo {Tipo} CUIT={Cuit}",
            documento.TipoDocumento, documento.Emisor.NumeroDocumento);

        // TODO: implementar integración real con AFIP WS-FE
        // 1. Autenticar via WSAA (token de acceso AFIP)
        // 2. Llamar WS-FE FECAESolicitar con datos del comprobante
        // 3. Parsear respuesta CAE/CAEA
        // 4. Retornar número CAE y fecha de vencimiento
        await Task.CompletedTask;

        _logger.LogWarning("FE Argentina: integración real con AFIP WS-FE aún no implementada. Retornando resultado simulado.");
        return FacturacionElectronicaResult.Exitoso($"CAE-{documento.Serie}-{documento.Numero}");
    }

    public async Task<FacturacionElectronicaResult> AnularDocumentoAsync(string idDocumento, string motivo, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("FE Argentina (AFIP): anulación no disponible en WS-FE standard. Id={Id}", idDocumento);
        // AFIP no permite anulación directa; se emite nota de crédito
        await Task.CompletedTask;
        return FacturacionElectronicaResult.Fallido("AFIP_NO_ANULACION", "AFIP no permite anulación directa. Emitir nota de crédito.");
    }

    public async Task<FacturacionElectronicaResult> ConsultarEstadoAsync(string idDocumento, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("FE Argentina (AFIP): consultando estado {Id}", idDocumento);
        await Task.CompletedTask;
        return FacturacionElectronicaResult.Exitoso(idDocumento);
    }
}

/// <summary>Opciones de configuración para FE Argentina (AFIP).</summary>
public sealed class ArgentinaFEOptions
{
    public const string SectionName = "FE:Argentina";
    public string EndpointWSAA { get; set; } = string.Empty;
    public string EndpointWSFE { get; set; } = string.Empty;
    public string Cuit { get; set; } = string.Empty;
    public string RutaCertificado { get; set; } = string.Empty;
    public string PasswordCertificado { get; set; } = string.Empty;
}
