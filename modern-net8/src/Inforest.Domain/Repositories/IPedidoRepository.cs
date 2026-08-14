using Inforest.Domain.Entities.Ventas;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Repositorio de pedidos — interfaz en Domain, implementación en Infrastructure.
/// Legacy: spIns_MPEDIDO, spUpd_MPEDIDO, MPEDIDO/DPEDIDO
/// </summary>
public interface IPedidoRepository
{
    Task<Pedido?> ObtenerPorCodigoAsync(string codigoPedido, CancellationToken cancellationToken = default);
    Task<IEnumerable<Pedido>> ObtenerAbiertosAsync(string codigoCaja, CancellationToken cancellationToken = default);
    Task InsertarAsync(Pedido pedido, CancellationToken cancellationToken = default);
    Task ActualizarAsync(Pedido pedido, CancellationToken cancellationToken = default);
}
