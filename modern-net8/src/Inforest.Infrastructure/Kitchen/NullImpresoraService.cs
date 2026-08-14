using Inforest.Application.Interfaces;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Implementación base de impresora mientras la integración hardware real sigue pendiente.
/// Legacy: impresoras térmicas y comandas por área.
/// Regla BR-011.
/// </summary>
internal sealed class NullImpresoraService : IImpresoraService
{
    private readonly ILogger<NullImpresoraService> _logger;

    public NullImpresoraService(ILogger<NullImpresoraService> logger)
    {
        _logger = logger;
    }

    public Task ImprimirTicketAsync(string contenido, string? nombreImpresora = null, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Impresión de ticket simulada en {Impresora}", nombreImpresora ?? "(default)");
        return Task.CompletedTask;
    }

    public Task ImprimirComandaAsync(string contenido, string codigoArea, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Impresión de comanda simulada para área {Area}", codigoArea);
        return Task.CompletedTask;
    }

    public Task AbrirCajonAsync(string? nombreImpresora = null, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Apertura de cajón simulada en {Impresora}", nombreImpresora ?? "(default)");
        return Task.CompletedTask;
    }
}
