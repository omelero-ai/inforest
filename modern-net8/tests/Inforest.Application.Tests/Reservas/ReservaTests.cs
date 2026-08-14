using Inforest.Application.Reservas;
using Inforest.Application.Turno;
using Inforest.Domain.Entities.Reservas;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Reservas;

// ──────────────────────────────────────────────────────────────────────────────
// Domain — entidad Reserva
// ──────────────────────────────────────────────────────────────────────────────

public class ReservaTests
{
    [Fact]
    public void Crear_DatosValidos_EstadoPendiente()
    {
        var r = Reserva.Crear("2600000001", DateTime.Today, "20:00", null, "Juan Perez", "999111222", "01", null, 4, "M01", null, DateTime.Today);
        Assert.Equal(EstadoReserva.Pendiente, r.Estado);
        Assert.Equal("2600000001", r.CodigoReserva);
        Assert.Equal(4, r.Pax);
    }

    [Fact]
    public void Crear_SinCliente_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            Reserva.Crear("2600000002", DateTime.Today, "20:00", null, null, null, null, null, 2, null, null, null));
        Assert.Equal("RESERVA_CLIENTE_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Crear_PaxCero_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            Reserva.Crear("2600000003", DateTime.Today, "20:00", null, "Cliente", null, null, null, 0, null, null, null));
        Assert.Equal("RESERVA_PAX_INVALIDO", ex.Code);
    }

    [Fact]
    public void Crear_CodigoVacio_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            Reserva.Crear("", DateTime.Today, "20:00", null, "Cliente", null, null, null, 2, null, null, null));
        Assert.Equal("RESERVA_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Anular_EstadoPendiente_CambiaEstado()
    {
        var r = Reserva.Crear("2600000004", DateTime.Today, "20:00", null, "Cliente", null, null, null, 2, null, null, null);
        r.Anular();
        Assert.Equal(EstadoReserva.Anulada, r.Estado);
    }

    [Fact]
    public void Anular_YaAnulada_LanzaExcepcion()
    {
        var r = Reserva.Crear("2600000005", DateTime.Today, "20:00", null, "Cliente", null, null, null, 2, null, null, null);
        r.Anular();
        var ex = Assert.Throws<DomainException>(() => r.Anular());
        Assert.Equal("RESERVA_YA_ANULADA", ex.Code);
    }

    [Fact]
    public void MarcarAtendida_EstadoPendiente_CambiaEstado()
    {
        var r = Reserva.Crear("2600000006", DateTime.Today, "20:00", null, "Cliente", null, null, null, 3, null, null, null);
        r.MarcarAtendida();
        Assert.Equal(EstadoReserva.Atendida, r.Estado);
    }

    [Fact]
    public void MarcarAtendida_YaAtendida_LanzaExcepcion()
    {
        var r = Reserva.Crear("2600000007", DateTime.Today, "20:00", null, "Cliente", null, null, null, 3, null, null, null);
        r.MarcarAtendida();
        var ex = Assert.Throws<DomainException>(() => r.MarcarAtendida());
        Assert.Equal("RESERVA_NO_PENDIENTE", ex.Code);
    }

    [Fact]
    public void Modificar_ReservaAtendida_LanzaExcepcion()
    {
        var r = Reserva.Crear("2600000008", DateTime.Today, "20:00", null, "Cliente", null, null, null, 3, null, null, null);
        r.MarcarAtendida();
        var ex = Assert.Throws<DomainException>(() =>
            r.Modificar(DateTime.Today, "21:00", null, "Otro", null, null, null, 2, null, null));
        Assert.Equal("RESERVA_YA_ATENDIDA", ex.Code);
    }

    [Fact]
    public void Modificar_ReservaAnulada_LanzaExcepcion()
    {
        var r = Reserva.Crear("2600000009", DateTime.Today, "20:00", null, "Cliente", null, null, null, 3, null, null, null);
        r.Anular();
        var ex = Assert.Throws<DomainException>(() =>
            r.Modificar(DateTime.Today, "21:00", null, "Otro", null, null, null, 2, null, null));
        Assert.Equal("RESERVA_YA_ANULADA", ex.Code);
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// Application — handlers
// ──────────────────────────────────────────────────────────────────────────────

public class CrearReservaHandlerTests
{
    private static Reserva BuildReserva(string codigo = "2600000001") =>
        Reserva.Crear(codigo, DateTime.Today, "20:00", null, "Juan Perez", "999", "01", null, 4, "M01", null, DateTime.Today);

    [Fact]
    public async Task Crear_DatosValidos_RetornaReservaConEstadoPendiente()
    {
        var repo = new Mock<IReservaRepository>();
        repo.Setup(r => r.GenerarProximoCodigoAsync(It.IsAny<CancellationToken>())).ReturnsAsync("2600000001");
        repo.Setup(r => r.InsertarAsync(It.IsAny<Reserva>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var diaContable = new Mock<IDiaContableService>();
        diaContable.Setup(s => s.ObtenerFechaServidorAsync(It.IsAny<CancellationToken>())).ReturnsAsync(DateTime.Today);

        var handler = new CrearReservaHandler(repo.Object, diaContable.Object);
        var result = await handler.HandleAsync(new CrearReservaCommand(DateTime.Today, "20:00", null, "Juan Perez", "999", "01", null, 4, "M01", null));

        Assert.True(result.EsExitoso);
        Assert.Equal(EstadoReserva.Pendiente, result.Valor!.Estado);
        repo.Verify(r => r.InsertarAsync(It.IsAny<Reserva>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Crear_SinCliente_RetornaError()
    {
        var repo = new Mock<IReservaRepository>();
        repo.Setup(r => r.GenerarProximoCodigoAsync(It.IsAny<CancellationToken>())).ReturnsAsync("2600000001");
        var diaContable = new Mock<IDiaContableService>();
        diaContable.Setup(s => s.ObtenerFechaServidorAsync(It.IsAny<CancellationToken>())).ReturnsAsync(DateTime.Today);

        var handler = new CrearReservaHandler(repo.Object, diaContable.Object);
        var result = await handler.HandleAsync(new CrearReservaCommand(DateTime.Today, "20:00", null, null, null, null, null, 2, null, null));

        Assert.False(result.EsExitoso);
        Assert.Equal("RESERVA_CLIENTE_REQUERIDO", result.CodigoError);
    }
}

public class AnularReservaHandlerTests
{
    [Fact]
    public async Task Anular_ReservaPendiente_RetornaOk()
    {
        var reserva = Reserva.Crear("2600000001", DateTime.Today, "20:00", null, "Cliente", null, null, null, 2, null, null, null);
        var repo = new Mock<IReservaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("2600000001", It.IsAny<CancellationToken>())).ReturnsAsync(reserva);
        repo.Setup(r => r.ActualizarAsync(It.IsAny<Reserva>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var handler = new AnularReservaHandler(repo.Object);
        var result = await handler.HandleAsync(new AnularReservaCommand("2600000001"));

        Assert.True(result.EsExitoso);
        Assert.Equal(EstadoReserva.Anulada, reserva.Estado);
    }

    [Fact]
    public async Task Anular_ReservaInexistente_RetornaError()
    {
        var repo = new Mock<IReservaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync(It.IsAny<string>(), It.IsAny<CancellationToken>())).ReturnsAsync((Reserva?)null);

        var handler = new AnularReservaHandler(repo.Object);
        var result = await handler.HandleAsync(new AnularReservaCommand("NOEXISTE"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RESERVA_NO_ENCONTRADA", result.CodigoError);
    }

    [Fact]
    public async Task Anular_ReservaYaAnulada_RetornaError()
    {
        var reserva = Reserva.Crear("2600000002", DateTime.Today, "20:00", null, "Cliente", null, null, null, 2, null, null, null);
        reserva.Anular();

        var repo = new Mock<IReservaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("2600000002", It.IsAny<CancellationToken>())).ReturnsAsync(reserva);

        var handler = new AnularReservaHandler(repo.Object);
        var result = await handler.HandleAsync(new AnularReservaCommand("2600000002"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RESERVA_YA_ANULADA", result.CodigoError);
    }
}

public class ModificarReservaHandlerTests
{
    [Fact]
    public async Task Modificar_ReservaPendiente_RetornaOk()
    {
        var reserva = Reserva.Crear("2600000001", DateTime.Today, "20:00", null, "Original", null, null, null, 2, null, null, null);
        var repo = new Mock<IReservaRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("2600000001", It.IsAny<CancellationToken>())).ReturnsAsync(reserva);
        repo.Setup(r => r.ActualizarAsync(It.IsAny<Reserva>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var handler = new ModificarReservaHandler(repo.Object);
        var result = await handler.HandleAsync(new ModificarReservaCommand("2600000001", DateTime.Today, "21:00", null, "Modificado", "888", null, null, 3, "M02", "obs"));

        Assert.True(result.EsExitoso);
        Assert.Equal("Modificado", reserva.NombreCliente);
        Assert.Equal(3, reserva.Pax);
        repo.Verify(r => r.ActualizarAsync(It.IsAny<Reserva>(), It.IsAny<CancellationToken>()), Times.Once);
    }
}

public class ObtenerReservasPorFechaHandlerTests
{
    [Fact]
    public async Task Obtener_SinReservas_RetornaListaVacia()
    {
        var repo = new Mock<IReservaRepository>();
        repo.Setup(r => r.ObtenerPorFechaAsync(It.IsAny<DateTime>(), It.IsAny<DateTime>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync([]);

        var handler = new ObtenerReservasPorFechaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerReservasPorFechaQuery(DateTime.Today, DateTime.Today));

        Assert.True(result.EsExitoso);
        Assert.Empty(result.Valor!);
    }

    [Fact]
    public async Task Obtener_ConReservas_RetornaResultados()
    {
        var lista = new List<Reserva>
        {
            Reserva.Crear("2600000001", DateTime.Today, "20:00", null, "Cliente A", null, null, null, 2, null, null, null),
            Reserva.Crear("2600000002", DateTime.Today, "21:00", null, "Cliente B", null, null, null, 4, null, null, null)
        };
        var repo = new Mock<IReservaRepository>();
        repo.Setup(r => r.ObtenerPorFechaAsync(It.IsAny<DateTime>(), It.IsAny<DateTime>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(lista);

        var handler = new ObtenerReservasPorFechaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerReservasPorFechaQuery(DateTime.Today, DateTime.Today));

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count());
    }
}
