using Inforest.Domain.Entities.Seguridad;

namespace Inforest.Application.Interfaces;

/// <summary>
/// RBAC en memoria basado en TUSUARIO/TGRUPOUSUARIO/TACCESO/TGRUPOACCESO.
/// ADR-006.
/// </summary>
public interface IRbacService
{
    Task<IReadOnlyCollection<PermisoAcceso>> ObtenerPermisosAsync(string codigoUsuario, string modulo, CancellationToken cancellationToken = default);

    Task<bool> TieneAccesoAsync(string codigoUsuario, string modulo, string codigoAcceso, CancellationToken cancellationToken = default);
}
