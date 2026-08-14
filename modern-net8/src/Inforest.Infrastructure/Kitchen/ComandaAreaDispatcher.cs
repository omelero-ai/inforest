using System.Text;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Cocina;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Adaptador base de comandas por área.
/// Legacy: ImprimePedido() + TAREAIMPRESORA.
/// Regla BR-011.
/// </summary>
internal sealed class ComandaAreaDispatcher : IComandaAreaDispatcher
{
    private readonly IImpresoraService _impresoraService;

    public ComandaAreaDispatcher(IImpresoraService impresoraService)
    {
        _impresoraService = impresoraService;
    }

    public async Task ImprimirAsync(OrdenProduccion orden, PlanProduccion plan, CancellationToken cancellationToken = default)
    {
        foreach (var comanda in plan.ComandasPorArea)
        {
            var contenido = FormatearComanda(orden, comanda);
            await _impresoraService.ImprimirComandaAsync(contenido, comanda.CodigoArea, cancellationToken);
        }
    }

    internal static string FormatearComanda(OrdenProduccion orden, ComandaAreaPlan comanda)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"PEDIDO: {orden.CodigoPedido}");
        sb.AppendLine($"AREA: {comanda.CodigoArea}");

        if (!string.IsNullOrWhiteSpace(orden.ObtenerNombreMesa()))
            sb.AppendLine($"MESA: {orden.ObtenerNombreMesa()}");

        if (!string.IsNullOrWhiteSpace(orden.NombreServidor))
            sb.AppendLine($"SERVIDOR: {orden.NombreServidor}");

        foreach (var item in comanda.Items)
        {
            sb.AppendLine($"{item.Cantidad:0.##} x {item.NombreSalida}");
            foreach (var condiment in item.Condimentos)
            {
                sb.AppendLine($"  - {condiment.Nombre}");
            }
        }

        return sb.ToString().TrimEnd();
    }
}
