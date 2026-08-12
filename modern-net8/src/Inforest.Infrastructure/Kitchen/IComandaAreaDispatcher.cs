using Inforest.Domain.Entities.Cocina;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Adaptador de impresión de comandas por área.
/// Legacy: ImprimePedido() + TAREAIMPRESORA.
/// Regla BR-011.
/// </summary>
internal interface IComandaAreaDispatcher
{
    Task ImprimirAsync(OrdenProduccion orden, PlanProduccion plan, CancellationToken cancellationToken = default);
}
