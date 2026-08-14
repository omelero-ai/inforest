using Inforest.Domain.Entities.Cocina;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Services;

/// <summary>
/// Router de producción por área e integración KDS.
/// Legacy: frmVenta.frm + modKDS.bas + TPRODUCTOAREA/TAREAIMPRESORA.
/// Reglas BR-010 y BR-011.
/// </summary>
public sealed class AreaProduccionRouter
{
    public PlanProduccion CrearPlan(
        OrdenProduccion orden,
        ConfiguracionProduccionCocina configuracion)
    {
        ArgumentNullException.ThrowIfNull(orden);
        ArgumentNullException.ThrowIfNull(configuracion);

        ValidarItems(orden.Items);

        var comandas = orden.Items
            .Where(item => item.ImprimeArea)
            .GroupBy(item => item.CodigoArea!, StringComparer.OrdinalIgnoreCase)
            .Select(group => new ComandaAreaPlan(group.Key, group.ToArray()))
            .OrderBy(group => group.CodigoArea, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        IReadOnlyList<KdsCanalPlan> canalesKds = [];

        if (configuracion.UsaKdsXmlExterno)
        {
            var itemsKds = orden.Items
                .Where(item => item.TieneDestinoKds)
                .ToArray();

            canalesKds = configuracion.TieneModeloSecundario
                ? itemsKds
                    .GroupBy(item => item.ObtenerCanalKds())
                    .Select(group => new KdsCanalPlan(group.Key, group.ToArray()))
                    .OrderBy(group => group.Canal)
                    .ToArray()
                : itemsKds.Length == 0
                    ? []
                    : [new KdsCanalPlan(0, itemsKds)];
        }

        return new PlanProduccion(
            comandas,
            canalesKds,
            configuracion.KdsHabilitado && configuracion.KdsInforestHabilitado);
    }

    private static void ValidarItems(IEnumerable<ItemProduccion> items)
    {
        foreach (var item in items)
        {
            if (item.ImprimeArea && string.IsNullOrWhiteSpace(item.CodigoArea))
            {
                throw new DomainException(
                    $"El producto {item.CodigoProducto} no tiene un área de producción configurada.",
                    "COCINA_AREA_REQUERIDA");
            }
        }
    }
}
