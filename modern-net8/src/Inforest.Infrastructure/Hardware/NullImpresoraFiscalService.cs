using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Stub de impresora fiscal Epson para ambientes sin hardware fiscal.
/// GAP: IFEpson.ocx es un OCX de 32 bits (Argentina, AFIP). Sin SDK .NET oficial.
/// Regla BR-PERIPH-005. ADR-009.
/// </summary>
internal sealed class NullImpresoraFiscalService : IImpresoraFiscalService
{
    private readonly ILogger<NullImpresoraFiscalService> _logger;

    public NullImpresoraFiscalService(ILogger<NullImpresoraFiscalService> logger) => _logger = logger;

    public Task<bool> ConfigurarAsync(ConfiguracionImpresoraFiscal configuracion, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("ImpresoraFiscal (Null): configuración simulada puerto={Puerto}", configuracion.Puerto);
        return Task.FromResult(true);
    }

    public Task<bool> EmitirFacturaAsync(FiscalDocumentoDto datosFactura, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("ImpresoraFiscal (Null): emisión de factura simulada letra={Letra} nro={Numero}",
            datosFactura.LetraComprobante, datosFactura.NumeroFactura);
        return Task.FromResult(true);
    }

    public Task<bool> CerrarJornadaAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("ImpresoraFiscal (Null): cierre de jornada fiscal simulado");
        return Task.FromResult(true);
    }
}
