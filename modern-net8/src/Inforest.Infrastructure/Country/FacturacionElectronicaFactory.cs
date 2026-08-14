using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Country;

/// <summary>
/// Fábrica que resuelve el gateway de FE según el código de país.
/// Legacy: selección de comportamiento fiscal por variable global `pais` en modDeclaracion.bas.
/// Regla BR-PERIPH-007. ADR-010.
/// </summary>
public sealed class FacturacionElectronicaFactory
{
    private readonly IReadOnlyDictionary<string, IFacturacionElectronicaGateway> _gateways;
    private readonly ILogger<FacturacionElectronicaFactory> _logger;

    public FacturacionElectronicaFactory(
        IEnumerable<IFacturacionElectronicaGateway> gateways,
        ILogger<FacturacionElectronicaFactory> logger)
    {
        _gateways = gateways.ToDictionary(g => g.CodigoPais, StringComparer.OrdinalIgnoreCase);
        _logger = logger;
    }

    /// <summary>
    /// Resuelve el gateway para el código de país indicado.
    /// Si no existe un gateway registrado para el país, retorna el NullFEGateway.
    /// </summary>
    public IFacturacionElectronicaGateway ObtenerGateway(string codigoPais)
    {
        if (_gateways.TryGetValue(codigoPais, out var gateway))
            return gateway;

        _logger.LogWarning("FacturacionElectronicaFactory: no se encontró gateway para país {CodigoPais}. Usando NullFEGateway.", codigoPais);
        return _gateways.TryGetValue("NULL", out var nullGateway)
            ? nullGateway
            : throw new InvalidOperationException("NullFEGateway no está registrado.");
    }
}
