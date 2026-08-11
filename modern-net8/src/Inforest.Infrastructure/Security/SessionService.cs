using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Seguridad;

namespace Inforest.Infrastructure.Security;

internal sealed class SessionService : ISessionService
{
    private readonly object _sync = new();

    public SesionOperativa? SesionActual { get; private set; }

    public Result<SesionOperativa> EstablecerSesion(SesionOperativa sesion)
    {
        ArgumentNullException.ThrowIfNull(sesion);

        lock (_sync)
        {
            if (SesionActual?.EstaActiva == true)
                return Result.Fail<SesionOperativa>("Ya existe una sesión activa.", "SEGURIDAD_SESION_ACTIVA");

            SesionActual = sesion;
            return Result.Ok(sesion);
        }
    }

    public Result CerrarSesion(DateTime fechaSalida)
    {
        lock (_sync)
        {
            if (SesionActual is null)
                return Result.Fail("No existe una sesión activa.", "SEGURIDAD_SESION_INEXISTENTE");

            var result = SesionActual.Cerrar(fechaSalida);
            if (result.EsExitoso)
                SesionActual = null;

            return result;
        }
    }
}
