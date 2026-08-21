using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Stub de biometría SecuGen para ambientes sin hardware biométrico.
/// GAP: sgfplibx.ocx es un OCX de 32 bits sin SDK .NET oficial. Se registra en known-gaps.
/// Regla BR-PERIPH-004. ADR-009.
/// </summary>
internal sealed class NullBiometriaService : IBiometriaService
{
    private const string MensajeNoDisponible =
        "Biometría no disponible (hardware SecuGen no conectado o SDK no instalado).";

    private readonly ILogger<NullBiometriaService> _logger;

    public NullBiometriaService(ILogger<NullBiometriaService> logger) => _logger = logger;

    public Task<BiometriaResult> CapturarHuellaAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Biometría (Null): captura de huella simulada");
        // Retorna template vacío — apto solo para desarrollo/test sin hardware
        return Task.FromResult(BiometriaResult.Exitoso(100, Array.Empty<byte>()));
    }

    public Task<BiometriaResult> VerificarUsuarioAsync(string codigoUsuario, byte[] templateHuella, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Biometría (Null): verificación simulada para usuario {CodigoUsuario}", codigoUsuario);
        // En modo Null, la verificación siempre es exitosa (para dev/test sin hardware real)
        return Task.FromResult(BiometriaResult.Exitoso(100));
    }

    public Task<BiometriaIdentificacionResult> IdentificarUsuarioAsync(string modulo, CancellationToken cancellationToken = default)
    {
        _logger.LogWarning(
            "Biometría (Null): identificación 1:N no disponible para módulo {Modulo}. " +
            "GAP-004: sgfplibx.ocx OCX 32-bit sin SDK .NET.", modulo);
        return Task.FromResult(BiometriaIdentificacionResult.Fallido(
            codigoError: -1,
            mensaje: MensajeNoDisponible));
    }
}
