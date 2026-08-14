using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Pedidos;

public interface IPedidoReadRepository
{
    Task<IReadOnlyList<Pedido>> ObtenerPorTurnoAsync(string codigoTurno, CancellationToken ct = default);
    Task<IReadOnlyList<Pedido>> ObtenerPorMesaAsync(string codigoMesa, CancellationToken ct = default);
    Task<IReadOnlyList<DetallePedido>> ObtenerDetalleAsync(string codigoPedido, CancellationToken ct = default);
}
