using Inforest.Domain.Entities.Cocina;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Adaptador de despacho KDS por XML/directorio.
/// Legacy: modKDS.bas.
/// Regla BR-010.
/// </summary>
internal interface IKdsDispatcher
{
    Task EnviarNuevaOrdenAsync(
        OrdenProduccion orden,
        PlanProduccion plan,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);

    Task EliminarOrdenAsync(
        OrdenProduccion orden,
        PlanProduccion plan,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);

    Task EliminarProductoAsync(
        OrdenProduccion orden,
        ItemProduccion item,
        PlanProduccion plan,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);

    Task<int> ProcesarBumpNotificationsAsync(
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);
}
