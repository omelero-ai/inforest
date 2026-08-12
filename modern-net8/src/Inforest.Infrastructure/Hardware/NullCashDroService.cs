using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Stub de CashDro para ambientes sin hardware.
/// Regla BR-PERIPH-002. ADR-011.
/// </summary>
internal sealed class NullCashDroService : ICashDroService
{
    private readonly ILogger<NullCashDroService> _logger;

    public NullCashDroService(ILogger<NullCashDroService> logger) => _logger = logger;

    public Task IniciarMotorAsync(string argumentos, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("CashDro (Null): motor simulado con argumentos={Argumentos}", argumentos);
        return Task.CompletedTask;
    }

    public Task AbrirCajonAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("CashDro (Null): apertura de cajón simulada");
        return Task.CompletedTask;
    }
}
