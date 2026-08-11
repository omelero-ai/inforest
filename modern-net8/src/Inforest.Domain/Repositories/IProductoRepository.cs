using Inforest.Domain.Entities.Productos;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Repositorio de productos — interfaz en Domain, implementación en Infrastructure.
/// Legacy: TPRODUCTO, TGRUPO, TSUBGRUPO
/// </summary>
public interface IProductoRepository
{
    Task<Producto?> ObtenerPorCodigoAsync(string codigoProducto, CancellationToken cancellationToken = default);
    Task<IEnumerable<Producto>> ObtenerActivosPorGrupoAsync(string codigoGrupo, CancellationToken cancellationToken = default);
    Task<IEnumerable<Producto>> BuscarAsync(string termino, CancellationToken cancellationToken = default);
}
