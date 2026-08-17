using Inforest.Domain.Common;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Pedidos;

/// <summary>
/// Obtiene el detalle completo de un pedido para presentación en FrmDetallePedido.
/// Legacy: frmDetallePedido.frm — Form_Load (vPedidoDetalle + vPedidoCombo + MPEDIDO cabecera).
/// Reglas: BR-PEDIDO-005, BR-PEDIDO-006.
/// </summary>
public sealed record ObtenerDetallePedidoExtendidoQuery(string CodigoPedido);

public sealed class ObtenerDetallePedidoExtendidoHandler
{
    private readonly IPedidoRepository _pedidoRepository;
    private readonly IPedidoReadRepository _readRepository;

    public ObtenerDetallePedidoExtendidoHandler(
        IPedidoRepository pedidoRepository,
        IPedidoReadRepository readRepository)
    {
        _pedidoRepository = pedidoRepository;
        _readRepository = readRepository;
    }

    public async Task<Result<DetallePedidoExtendidoResultado>> HandleAsync(
        ObtenerDetallePedidoExtendidoQuery query,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoPedido))
            return Result.Fail<DetallePedidoExtendidoResultado>("El código de pedido es requerido.", "PEDIDO_CODIGO_REQUERIDO");

        // Cabecera — IPedidoRepository (write side, tiene ObtenerPorCodigoAsync)
        var pedido = await _pedidoRepository.ObtenerPorCodigoAsync(query.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail<DetallePedidoExtendidoResultado>("Pedido no encontrado.", "PEDIDO_NO_ENCONTRADO");

        // Items enriquecidos y combos — queries independientes, lanzadas en paralelo
        var itemsTask  = _readRepository.ObtenerDetalleExtendidoAsync(query.CodigoPedido, ct);
        var combosTask = _readRepository.ObtenerCombosAsync(query.CodigoPedido, ct);
        await Task.WhenAll(itemsTask, combosTask);
        var items  = itemsTask.Result;
        var combos = combosTask.Result;

        // Legacy: total = sum(nVenta) from DPEDIDO WHERE tEstadoItem='N'
        var total = items.Sum(i => i.SubTotal);

        var resultado = new DetallePedidoExtendidoResultado(
            CodigoPedido: pedido.CodigoPedido,
            NumeroAdultos: pedido.NumeroPersonas,
            NumeroNinos: 0,          // MPEDIDO.nNino — no expuesto en entidad Pedido aún (GAP menor)
            FechaRegistro: pedido.FechaRegistro,
            Observacion: pedido.Observacion,
            ClienteDelivery: null,   // TDELIVERY.tNombre — fuera del scope de este corte
            Total: total,
            Items: items,
            Combos: combos);

        return Result.Ok(resultado);
    }
}
