namespace Inforest.Domain.Entities.Cocina;

/// <summary>
/// Plan resultante de routing de producción por área y por canal KDS.
/// Legacy: frmVenta.frm + modKDS.bas.
/// Reglas BR-010 y BR-011.
/// </summary>
public sealed class PlanProduccion
{
    public IReadOnlyList<ComandaAreaPlan> ComandasPorArea { get; }
    public IReadOnlyList<KdsCanalPlan> CanalesKds { get; }
    public bool RequiereSincronizacionKdsInforest { get; }

    public PlanProduccion(
        IReadOnlyList<ComandaAreaPlan> comandasPorArea,
        IReadOnlyList<KdsCanalPlan> canalesKds,
        bool requiereSincronizacionKdsInforest)
    {
        ComandasPorArea = comandasPorArea;
        CanalesKds = canalesKds;
        RequiereSincronizacionKdsInforest = requiereSincronizacionKdsInforest;
    }
}

/// <summary>
/// Agrupación de ítems a imprimir en una comanda por área.
/// </summary>
public sealed record ComandaAreaPlan(string CodigoArea, IReadOnlyList<ItemProduccion> Items);

/// <summary>
/// Agrupación de ítems a despachar a un canal/modelo KDS.
/// </summary>
public sealed record KdsCanalPlan(int Canal, IReadOnlyList<ItemProduccion> Items);
