using Inforest.Application.Kitchen;
using Inforest.Domain.Entities.Cocina;
using Moq;

namespace Inforest.Application.Tests.Kitchen;

public sealed class MensajeCocinaHandlerTests
{
    [Fact]
    public async Task ObtenerMensajes_RangoInvalido_RetornaFallo()
    {
        var repo = new Mock<IMensajeCocinaRepository>();
        var handler = new ObtenerMensajesCocinaHandler(repo.Object);

        var result = await handler.HandleAsync(new ObtenerMensajesCocinaQuery(DateTime.Today, DateTime.Today.AddDays(-1)));

        Assert.False(result.EsExitoso);
        Assert.Equal("MENSAJE_COCINA_RANGO_INVALIDO", result.CodigoError);
    }

    [Fact]
    public async Task AgregarMensaje_ActivoYMaximoAlcanzado_RetornaFallo()
    {
        var repo = new Mock<IMensajeCocinaRepository>();
        repo.Setup(r => r.ContarActivosAsync(null, default)).ReturnsAsync(MensajeCocina.MaximoMensajesActivos);
        var handler = new AgregarMensajeCocinaHandler(repo.Object);

        var result = await handler.HandleAsync(new AgregarMensajeCocinaCommand("mensaje", "USR1", "C01", true, false));

        Assert.False(result.EsExitoso);
        Assert.Equal("MENSAJE_COCINA_MAXIMO_ACTIVOS", result.CodigoError);
    }

    [Fact]
    public async Task AgregarMensaje_Normal_GeneraCodigoYGuarda()
    {
        var repo = new Mock<IMensajeCocinaRepository>();
        repo.Setup(r => r.ContarActivosAsync(null, default)).ReturnsAsync(2);
        repo.Setup(r => r.ObtenerProximoCodigoAsync(It.IsAny<DateTime>(), default)).ReturnsAsync("26000007");
        repo.Setup(r => r.AgregarAsync(It.IsAny<MensajeCocina>(), default)).Returns(Task.CompletedTask);

        MensajeCocina? capturado = null;
        repo.Setup(r => r.AgregarAsync(It.IsAny<MensajeCocina>(), default))
            .Callback<MensajeCocina, CancellationToken>((m, _) => capturado = m)
            .Returns(Task.CompletedTask);

        var handler = new AgregarMensajeCocinaHandler(repo.Object);
        var result = await handler.HandleAsync(new AgregarMensajeCocinaCommand("mensaje nuevo", "USR1", "C01", true, true));

        Assert.True(result.EsExitoso);
        Assert.Equal("26000007", result.Valor);
        Assert.NotNull(capturado);
        Assert.Equal("MENSAJE NUEVO", capturado!.Mensaje);
        Assert.True(capturado.ActivarAlerta);
    }

    [Fact]
    public async Task ModificarMensaje_NoExiste_RetornaFallo()
    {
        var repo = new Mock<IMensajeCocinaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("26000001", default)).ReturnsAsync((MensajeCocina?)null);
        var handler = new ModificarMensajeCocinaHandler(repo.Object);

        var result = await handler.HandleAsync(new ModificarMensajeCocinaCommand("26000001", "mensaje", "USR1", "C01", true, false));

        Assert.False(result.EsExitoso);
        Assert.Equal("MENSAJE_COCINA_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task ModificarMensaje_ActivoYLimiteAlcanzado_RetornaFallo()
    {
        var existente = MensajeCocina.Crear("26000001", "USR1", "mensaje", "C01", false, false);
        var repo = new Mock<IMensajeCocinaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("26000001", default)).ReturnsAsync(existente);
        repo.Setup(r => r.ContarActivosAsync("26000001", default)).ReturnsAsync(MensajeCocina.MaximoMensajesActivos);
        var handler = new ModificarMensajeCocinaHandler(repo.Object);

        var result = await handler.HandleAsync(new ModificarMensajeCocinaCommand("26000001", "mensaje", "USR1", "C01", true, false));

        Assert.False(result.EsExitoso);
        Assert.Equal("MENSAJE_COCINA_MAXIMO_ACTIVOS", result.CodigoError);
    }

    [Fact]
    public async Task ModificarMensaje_Existente_Actualiza()
    {
        var existente = MensajeCocina.Crear("26000001", "USR1", "mensaje", "C01", true, false);
        var repo = new Mock<IMensajeCocinaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("26000001", default)).ReturnsAsync(existente);
        repo.Setup(r => r.ContarActivosAsync("26000001", default)).ReturnsAsync(1);
        repo.Setup(r => r.ModificarAsync(existente, default)).Returns(Task.CompletedTask);

        var handler = new ModificarMensajeCocinaHandler(repo.Object);
        var result = await handler.HandleAsync(new ModificarMensajeCocinaCommand("26000001", "nuevo", "USR2", "C02", false, true));

        Assert.True(result.EsExitoso);
        Assert.Equal("NUEVO", existente.Mensaje);
        Assert.Equal("USR2", existente.UsuarioModificacion);
        Assert.False(existente.Activo);
    }

    [Fact]
    public async Task EliminarMensaje_NoExiste_RetornaFallo()
    {
        var repo = new Mock<IMensajeCocinaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("26000001", default)).ReturnsAsync((MensajeCocina?)null);
        var handler = new EliminarMensajeCocinaHandler(repo.Object);

        var result = await handler.HandleAsync(new EliminarMensajeCocinaCommand("26000001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("MENSAJE_COCINA_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task EliminarMensaje_Existe_Elimina()
    {
        var existente = MensajeCocina.Crear("26000001", "USR1", "mensaje", "C01", true, false);
        var repo = new Mock<IMensajeCocinaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("26000001", default)).ReturnsAsync(existente);
        repo.Setup(r => r.EliminarAsync("26000001", default)).Returns(Task.CompletedTask);
        var handler = new EliminarMensajeCocinaHandler(repo.Object);

        var result = await handler.HandleAsync(new EliminarMensajeCocinaCommand("26000001"));

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.EliminarAsync("26000001", default), Times.Once);
    }

    [Fact]
    public async Task CerrarMensajesPorCaja_UsuarioVacio_RetornaFallo()
    {
        var repo = new Mock<IMensajeCocinaRepository>();
        var handler = new CerrarMensajesCocinaPorCajaHandler(repo.Object);

        var result = await handler.HandleAsync(new CerrarMensajesCocinaPorCajaCommand("", "C01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("MENSAJE_COCINA_USUARIO_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task CerrarMensajesPorCaja_Normal_EjecutaStoredProcedure()
    {
        var repo = new Mock<IMensajeCocinaRepository>();
        repo.Setup(r => r.CerrarActivosPorCajaAsync("USR1", "C01", default)).Returns(Task.CompletedTask);
        var handler = new CerrarMensajesCocinaPorCajaHandler(repo.Object);

        var result = await handler.HandleAsync(new CerrarMensajesCocinaPorCajaCommand("USR1", "C01"));

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.CerrarActivosPorCajaAsync("USR1", "C01", default), Times.Once);
    }
}
