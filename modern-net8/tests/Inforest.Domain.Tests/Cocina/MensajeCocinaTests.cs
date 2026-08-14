using Inforest.Domain.Entities.Cocina;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Cocina;

public sealed class MensajeCocinaTests
{
    [Fact]
    public void Crear_NormalizaMensajeAMayusculas()
    {
        var mensaje = MensajeCocina.Crear("26000001", "USR1", "mensaje cocina", "C01", true, false);

        Assert.Equal("MENSAJE COCINA", mensaje.Mensaje);
        Assert.Equal("USR1", mensaje.UsuarioRegistro);
    }

    [Fact]
    public void Crear_MensajeVacio_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() => MensajeCocina.Crear("26000001", "USR1", "", "C01", true, false));

        Assert.Equal("MENSAJE_COCINA_REQUERIDO", ex.Code);
    }

    [Theory]
    [InlineData("contiene &")]
    [InlineData("contiene '")]
    [InlineData("contiene \"")]
    public void Crear_CaracterInvalido_LanzaExcepcion(string valor)
    {
        var ex = Assert.Throws<DomainException>(() => MensajeCocina.Crear("26000001", "USR1", valor, "C01", true, false));

        Assert.Equal("MENSAJE_COCINA_CARACTER_INVALIDO", ex.Code);
    }

    [Fact]
    public void Crear_LongitudMayor95_LanzaExcepcion()
    {
        var texto = new string('A', MensajeCocina.LongitudMaximaMensaje + 1);
        var ex = Assert.Throws<DomainException>(() => MensajeCocina.Crear("26000001", "USR1", texto, "C01", true, false));

        Assert.Equal("MENSAJE_COCINA_LONGITUD_INVALIDA", ex.Code);
    }

    [Fact]
    public void Actualizar_ModificaEstadoYAlerta()
    {
        var mensaje = MensajeCocina.Crear("26000001", "USR1", "mensaje", "C01", true, false);

        mensaje.Actualizar("USR2", "nuevo mensaje", "C02", false, true);

        Assert.Equal("NUEVO MENSAJE", mensaje.Mensaje);
        Assert.Equal("USR2", mensaje.UsuarioModificacion);
        Assert.Equal("C02", mensaje.Caja);
        Assert.False(mensaje.Activo);
        Assert.True(mensaje.ActivarAlerta);
    }
}
