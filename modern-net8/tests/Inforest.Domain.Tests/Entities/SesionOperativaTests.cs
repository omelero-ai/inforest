using Inforest.Domain.Entities.Seguridad;

namespace Inforest.Domain.Tests.Entities;

public sealed class SesionOperativaTests
{
    [Fact]
    public void Iniciar_CreaSesionActivaConPermisos()
    {
        var permisos = new[]
        {
            new PermisoAcceso("ACC001", "01", "frmCaja", "btnCobrar", "Cobrar")
        };

        var sesion = SesionOperativa.Iniciar(
            "USR01",
            "Administrador",
            "01",
            "01",
            "CA1",
            "POS-01",
            "INFOREST",
            15,
            new DateTime(2026, 8, 11, 12, 0, 0, DateTimeKind.Utc),
            permisos);

        Assert.True(sesion.EstaActiva);
        Assert.True(sesion.TienePermiso("ACC001"));
        Assert.Equal("USR01", sesion.CodigoUsuario);
    }

    [Fact]
    public void Cerrar_SegundaVezRetornaError()
    {
        var sesion = SesionOperativa.Iniciar(
            "USR01",
            "Administrador",
            "01",
            "01",
            "CA1",
            "POS-01",
            "INFOREST",
            15,
            DateTime.UtcNow,
            Array.Empty<PermisoAcceso>());

        var primerCierre = sesion.Cerrar(DateTime.UtcNow);
        var segundoCierre = sesion.Cerrar(DateTime.UtcNow);

        Assert.True(primerCierre.EsExitoso);
        Assert.False(segundoCierre.EsExitoso);
        Assert.Equal("SEGURIDAD_SESION_CERRADA", segundoCierre.CodigoError);
    }
}
