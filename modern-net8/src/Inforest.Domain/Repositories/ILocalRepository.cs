using Inforest.Domain.Entities.Delivery;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Repositorio de locales del restaurante (multi-local).
/// <para>
/// Legacy: tabla <c>TLOCAL</c>.
/// Controla IP, base de datos INFOREST por local y participación en réplica.
/// </para>
/// Regla BR-DEL-007.
/// </summary>
public interface ILocalRepository
{
    Task<LocalRestaurante?> ObtenerPorCodigoAsync(string codigoLocal, CancellationToken cancellationToken = default);
    Task<IEnumerable<LocalRestaurante>> ObtenerActivosAsync(CancellationToken cancellationToken = default);
    Task<IEnumerable<LocalRestaurante>> ObtenerConReplicaAsync(CancellationToken cancellationToken = default);
}
