using Inforest.Domain.Common;
using Inforest.Domain.Entities.Cocina;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Servicio de aplicación para routing de producción, comandas y KDS.
/// Legacy: frmVenta.frm + frmCajaRapida.frm + modKDS.bas.
/// Reglas BR-010 y BR-011.
/// </summary>
public interface IProduccionCocinaService
{
    PlanProduccion CrearPlan(OrdenProduccion orden, ConfiguracionProduccionCocina configuracion);

    Task<Result> EnviarNuevaOrdenAsync(
        OrdenProduccion orden,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);

    Task<Result> ReenviarOrdenAsync(
        OrdenProduccion orden,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);

    Task<Result> EliminarOrdenAsync(
        OrdenProduccion orden,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);

    Task<Result> EliminarProductoAsync(
        OrdenProduccion orden,
        ItemProduccion item,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);

    Task<int> ProcesarBumpNotificationsAsync(
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default);
}
