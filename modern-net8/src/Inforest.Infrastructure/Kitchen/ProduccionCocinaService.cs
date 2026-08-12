using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Cocina;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Services;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Orquestador de producción/cocina equivalente al flujo operativo Legacy.
/// Legacy: frmVenta.frm + modKDS.bas + impresión de áreas.
/// Reglas BR-010 y BR-011.
/// </summary>
internal sealed class ProduccionCocinaService : IProduccionCocinaService
{
    private readonly AreaProduccionRouter _router;
    private readonly IComandaAreaDispatcher _comandaDispatcher;
    private readonly IKdsDispatcher _kdsDispatcher;
    private readonly IKdsLegacyGateway _legacyGateway;

    public ProduccionCocinaService(
        AreaProduccionRouter router,
        IComandaAreaDispatcher comandaDispatcher,
        IKdsDispatcher kdsDispatcher,
        IKdsLegacyGateway legacyGateway)
    {
        _router = router;
        _comandaDispatcher = comandaDispatcher;
        _kdsDispatcher = kdsDispatcher;
        _legacyGateway = legacyGateway;
    }

    public PlanProduccion CrearPlan(OrdenProduccion orden, ConfiguracionProduccionCocina configuracion)
        => _router.CrearPlan(orden, configuracion);

    public Task<Result> ReenviarOrdenAsync(
        OrdenProduccion orden,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
        => EnviarNuevaOrdenAsync(orden, configuracion, cancellationToken);

    public async Task<Result> EnviarNuevaOrdenAsync(
        OrdenProduccion orden,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
    {
        return await EjecutarAsync(
            orden,
            configuracion,
            async plan =>
            {
                await _comandaDispatcher.ImprimirAsync(orden, plan, cancellationToken);

                if (configuracion.UsaKdsXmlExterno)
                    await _kdsDispatcher.EnviarNuevaOrdenAsync(orden, plan, configuracion, cancellationToken);

                if (plan.RequiereSincronizacionKdsInforest)
                    await _legacyGateway.SincronizarOrdenInforestAsync(
                        orden.CodigoPedido,
                        configuracion.KdsPorAreaHabilitado ? 2 : 1,
                        cancellationToken);
            });
    }

    public async Task<Result> EliminarOrdenAsync(
        OrdenProduccion orden,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
    {
        return await EjecutarAsync(
            orden,
            configuracion,
            async plan =>
            {
                if (configuracion.UsaKdsXmlExterno)
                    await _kdsDispatcher.EliminarOrdenAsync(orden, plan, configuracion, cancellationToken);

                if (plan.RequiereSincronizacionKdsInforest)
                    await _legacyGateway.SincronizarOrdenInforestAsync(
                        orden.CodigoPedido,
                        configuracion.KdsPorAreaHabilitado ? 2 : 1,
                        cancellationToken);
            });
    }

    public async Task<Result> EliminarProductoAsync(
        OrdenProduccion orden,
        ItemProduccion item,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
    {
        return await EjecutarAsync(
            orden,
            configuracion,
            async plan =>
            {
                if (configuracion.UsaKdsXmlExterno)
                    await _kdsDispatcher.EliminarProductoAsync(orden, item, plan, configuracion, cancellationToken);

                if (plan.RequiereSincronizacionKdsInforest)
                    await _legacyGateway.SincronizarOrdenInforestAsync(
                        orden.CodigoPedido,
                        configuracion.KdsPorAreaHabilitado ? 2 : 1,
                        cancellationToken);
            });
    }

    public Task<int> ProcesarBumpNotificationsAsync(
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
        => _kdsDispatcher.ProcesarBumpNotificationsAsync(configuracion, cancellationToken);

    private async Task<Result> EjecutarAsync(
        OrdenProduccion orden,
        ConfiguracionProduccionCocina configuracion,
        Func<PlanProduccion, Task> operation)
    {
        try
        {
            var plan = _router.CrearPlan(orden, configuracion);
            await operation(plan);
            return Result.Ok();
        }
        catch (DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
        catch (Exception ex)
        {
            return Result.Fail(ex.Message, "COCINA_PRODUCCION_ERROR");
        }
    }
}
