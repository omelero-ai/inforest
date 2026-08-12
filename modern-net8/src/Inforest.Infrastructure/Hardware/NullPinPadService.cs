using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Stub de PinPad para desarrollo y ambientes sin hardware.
/// Regla BR-PERIPH-001. ADR-009.
/// </summary>
internal sealed class NullPinPadService : IPinPadService
{
    private readonly ILogger<NullPinPadService> _logger;

    public NullPinPadService(ILogger<NullPinPadService> logger) => _logger = logger;

    public Task<PinPadResult> InicializarAsync(string configPath, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad (Null): inicialización simulada con {ConfigPath}", configPath);
        return Task.FromResult(PinPadResult.Ok("Simulado"));
    }

    public Task<PinPadResult> EjecutarVentaAsync(decimal monto, int cuotas, int timeoutSegundos = 60, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad (Null): venta simulada monto={Monto} cuotas={Cuotas}", monto, cuotas);
        return Task.FromResult(PinPadResult.Ok("Venta simulada", "AUTH-000000", "SIM-001"));
    }

    public Task<PinPadResult> AnularAsync(string numeroOperacion, int timeoutSegundos = 60, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad (Null): anulación simulada op={NumeroOperacion}", numeroOperacion);
        return Task.FromResult(PinPadResult.Ok("Anulación simulada"));
    }

    public Task<PinPadResult> CerrarAsync(int timeoutSegundos = 60, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad (Null): cierre de lote simulado");
        return Task.FromResult(PinPadResult.Ok("Cierre simulado"));
    }

    public Task<PinPadResult> TestComunicacionAsync(int timeoutSegundos = 30, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad (Null): test de comunicación simulado");
        return Task.FromResult(PinPadResult.Ok("Comunicación OK (simulado)"));
    }

    public Task<PinPadResult> SimularCuotasAsync(decimal monto, int cuotas, int timeoutSegundos = 30, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad (Null): simulación de cuotas monto={Monto} cuotas={Cuotas}", monto, cuotas);
        return Task.FromResult(PinPadResult.Ok("Simulación de cuotas OK (simulado)"));
    }

    public Task CerrarPuertoAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad (Null): cierre de puerto simulado");
        return Task.CompletedTask;
    }
}
