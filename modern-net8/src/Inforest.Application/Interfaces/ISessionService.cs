using Inforest.Domain.Common;
using Inforest.Domain.Entities.Seguridad;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Contexto de sesión actual del cliente desktop.
/// Reemplaza variables globales + USUARIO.INI.
/// </summary>
public interface ISessionService
{
    SesionOperativa? SesionActual { get; }

    Result<SesionOperativa> EstablecerSesion(SesionOperativa sesion);

    Result CerrarSesion(DateTime fechaSalida);
}
