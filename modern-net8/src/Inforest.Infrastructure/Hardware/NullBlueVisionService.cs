using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Stub de BlueVision para ambientes sin servidor BlueVision.
/// Regla BR-PERIPH-003. ADR-012.
/// </summary>
internal sealed class NullBlueVisionService : IBlueVisionService
{
    private readonly ILogger<NullBlueVisionService> _logger;

    public NullBlueVisionService(ILogger<NullBlueVisionService> logger) => _logger = logger;

    public Task<bool> EnviarTicketAsync(BlueVisionTicketDto ticket, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("BlueVision (Null): ticket {TicketId} simulado", ticket.Id);
        return Task.FromResult(true);
    }
}
