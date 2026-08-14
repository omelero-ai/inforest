using Inforest.Domain.Entities.Configuracion;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Turno;

public class TurnoTests
{
    [Fact]
    public void AbrirTurnoValido_CreaTurnoAbierto()
    {
        var turno = Inforest.Domain.Entities.Configuracion.Turno.Abrir("2401010001", "CA1", "USR1", new DateTime(2026, 8, 12), 100m);

        Assert.Equal("2401010001", turno.CodigoTurno);
        Assert.False(turno.Cerrado);
        Assert.Equal(100m, turno.MontoInicial);
    }

    [Fact]
    public void CerrarTurnoYaCerrado_LanzaDomainException()
    {
        var turno = Inforest.Domain.Entities.Configuracion.Turno.Abrir("2401010001", "CA1", "USR1", new DateTime(2026, 8, 12), 100m);
        turno.Cerrar(250m);

        var ex = Assert.Throws<DomainException>(() => turno.Cerrar(300m));
        Assert.Equal("TURNO_YA_CERRADO", ex.Code);
    }

    [Fact]
    public void DiaContable_AbrirYCerrar_CambiaEstadoCorrectamente()
    {
        var diaContable = DiaContable.Crear(new DateTime(2026, 8, 12), "USR1");
        diaContable.Cerrar("ADMIN1");

        Assert.True(diaContable.Apertura);
        Assert.True(diaContable.Cierre);
        Assert.Equal("ADMIN1", diaContable.UsuarioCierre);
        Assert.NotNull(diaContable.FechaRegistroCierre);
    }
}
