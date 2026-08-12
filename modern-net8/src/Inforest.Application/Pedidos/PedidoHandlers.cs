using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Pedidos;

public sealed record ObtenerPedidoPorCodigoQuery(string CodigoPedido);

public sealed class ObtenerPedidoPorCodigoHandler
{
    private readonly IPedidoRepository _pedidoRepository;

    public ObtenerPedidoPorCodigoHandler(IPedidoRepository pedidoRepository)
        => _pedidoRepository = pedidoRepository;

    public async Task<Result<Pedido?>> HandleAsync(ObtenerPedidoPorCodigoQuery query, CancellationToken ct = default)
        => Result.Ok<Pedido?>(await _pedidoRepository.ObtenerPorCodigoAsync(query.CodigoPedido, ct));
}

public sealed record ObtenerPedidosPorTurnoQuery(string CodigoTurno);

public sealed class ObtenerPedidosPorTurnoHandler
{
    private readonly IPedidoReadRepository _pedidoRepository;

    public ObtenerPedidosPorTurnoHandler(IPedidoReadRepository pedidoRepository)
        => _pedidoRepository = pedidoRepository;

    public async Task<Result<IReadOnlyList<Pedido>>> HandleAsync(ObtenerPedidosPorTurnoQuery query, CancellationToken ct = default)
        => Result.Ok(await _pedidoRepository.ObtenerPorTurnoAsync(query.CodigoTurno, ct));
}

public sealed record ObtenerPedidosPorMesaQuery(string CodigoMesa);

public sealed class ObtenerPedidosPorMesaHandler
{
    private readonly IPedidoReadRepository _pedidoRepository;

    public ObtenerPedidosPorMesaHandler(IPedidoReadRepository pedidoRepository)
        => _pedidoRepository = pedidoRepository;

    public async Task<Result<IReadOnlyList<Pedido>>> HandleAsync(ObtenerPedidosPorMesaQuery query, CancellationToken ct = default)
        => Result.Ok(await _pedidoRepository.ObtenerPorMesaAsync(query.CodigoMesa, ct));
}
