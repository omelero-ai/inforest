using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Country;

/// <summary>
/// Stub de Facturación Electrónica para países sin FE configurada o durante desarrollo.
/// Regla BR-PERIPH-006. ADR-010.
/// </summary>
internal sealed class NullFEGateway : IFacturacionElectronicaGateway
{
    public string CodigoPais => "NULL";
    private readonly ILogger<NullFEGateway> _logger;

    public NullFEGateway(ILogger<NullFEGateway> logger) => _logger = logger;

    public Task<FacturacionElectronicaResult> EmitirDocumentoAsync(DocumentoFEDto documento, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("FE (Null): emisión simulada {Tipo} {Serie}-{Numero}", documento.TipoDocumento, documento.Serie, documento.Numero);
        return Task.FromResult(FacturacionElectronicaResult.Exitoso("SIM-000000000"));
    }

    public Task<FacturacionElectronicaResult> AnularDocumentoAsync(string idDocumento, string motivo, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("FE (Null): anulación simulada {Id}", idDocumento);
        return Task.FromResult(FacturacionElectronicaResult.Exitoso(null));
    }

    public Task<FacturacionElectronicaResult> ConsultarEstadoAsync(string idDocumento, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("FE (Null): consulta de estado simulada {Id}", idDocumento);
        return Task.FromResult(FacturacionElectronicaResult.Exitoso(null));
    }
}
