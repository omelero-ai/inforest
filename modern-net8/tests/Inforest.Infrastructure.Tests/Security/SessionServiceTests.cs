using Inforest.Domain.Entities.Seguridad;
using Inforest.Infrastructure.Security;

namespace Inforest.Infrastructure.Tests.Security;

public sealed class SessionServiceTests
{
    [Fact]
    public void EstablecerSesion_DosVeces_SoloPermiteUnaActiva()
    {
        var service = new SessionService();
        var sesion1 = CreateSession("USR01");
        var sesion2 = CreateSession("USR02");

        var first = service.EstablecerSesion(sesion1);
        var second = service.EstablecerSesion(sesion2);

        Assert.True(first.EsExitoso);
        Assert.False(second.EsExitoso);
        Assert.Equal("USR01", service.SesionActual?.CodigoUsuario);
    }

    [Fact]
    public void CerrarSesion_LimpiaSesionActual()
    {
        var service = new SessionService();
        service.EstablecerSesion(CreateSession("USR01"));

        var result = service.CerrarSesion(DateTime.UtcNow);

        Assert.True(result.EsExitoso);
        Assert.Null(service.SesionActual);
    }

    private static SesionOperativa CreateSession(string codigoUsuario)
        => SesionOperativa.Iniciar(
            codigoUsuario,
            "Nombre",
            "01",
            "01",
            "CA1",
            "POS-01",
            "INFOREST",
            1,
            DateTime.UtcNow,
            Array.Empty<PermisoAcceso>());
}
