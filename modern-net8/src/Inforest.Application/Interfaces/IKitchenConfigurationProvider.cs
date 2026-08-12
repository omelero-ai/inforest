using Inforest.Domain.Entities.Cocina;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Proveedor de configuración de cocina/comanda/KDS.
/// Legacy: TPARAMETRO + TAREAIMPRESORA + TPRODUCTOAREA.
/// Reglas BR-010 y BR-011.
/// </summary>
public interface IKitchenConfigurationProvider
{
    Task<ConfiguracionProduccionCocina> ObtenerConfiguracionAsync(string codigoCaja, CancellationToken cancellationToken = default);

    Task<IReadOnlyDictionary<string, string?>> ObtenerImpresorasPorAreaAsync(string codigoCaja, CancellationToken cancellationToken = default);

    Task<IReadOnlyCollection<string>> ObtenerAreasProductoAsync(string codigoProducto, CancellationToken cancellationToken = default);
}
