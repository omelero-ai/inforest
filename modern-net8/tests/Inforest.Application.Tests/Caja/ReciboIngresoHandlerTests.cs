using Inforest.Application.Caja;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Caja;

/// <summary>
/// Tests de handlers para recibos de ingreso.
/// Cubre BR-RECIBO-001, BR-RECIBO-003, BR-RECIBO-006, BR-RECIBO-011.
/// </summary>
public class ReciboIngresoHandlerTests
{
    // ── ObtenerIngresosHandler ───────────────────────────────────────────────

    [Fact]
    public async Task ObtenerIngresos_RangoValido_RetornaLista()
    {
        var repo = new Mock<IReciboIngresoRepository>();
        var lista = new List<ReciboIngreso>
        {
            ReciboIngreso.Registrar("2600000001", DateTime.Today, 100m, "Test")
        };
        repo.Setup(r => r.ObtenerPorFechaAsync(It.IsAny<DateTime>(), It.IsAny<DateTime>(), null, default))
            .ReturnsAsync(lista.AsReadOnly());

        var handler = new ObtenerIngresosHandler(repo.Object);
        var result  = await handler.HandleAsync(new ObtenerIngresosQuery(DateTime.Today.AddDays(-1), DateTime.Today));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor);
    }

    [Fact]
    public async Task ObtenerIngresos_FechaDesdeEsMayor_RetornaFallo()
    {
        var repo    = new Mock<IReciboIngresoRepository>();
        var handler = new ObtenerIngresosHandler(repo.Object);
        var result  = await handler.HandleAsync(
            new ObtenerIngresosQuery(DateTime.Today.AddDays(1), DateTime.Today));

        Assert.False(result.EsExitoso);
        Assert.Equal("RECIBO_RANGO_FECHAS_INVALIDO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerIngresos_ConUsuarioFiltro_PasaFiltroAlRepo()
    {
        // BR-RECIBO-006: lMCPV
        var repo = new Mock<IReciboIngresoRepository>();
        repo.Setup(r => r.ObtenerPorFechaAsync(It.IsAny<DateTime>(), It.IsAny<DateTime>(), "USER01", default))
            .ReturnsAsync(new List<ReciboIngreso>().AsReadOnly());

        var handler = new ObtenerIngresosHandler(repo.Object);
        await handler.HandleAsync(new ObtenerIngresosQuery(DateTime.Today, DateTime.Today, "USER01"));

        repo.Verify(r => r.ObtenerPorFechaAsync(It.IsAny<DateTime>(), It.IsAny<DateTime>(), "USER01", default), Times.Once);
    }

    // ── RegistrarIngresoHandler ──────────────────────────────────────────────

    [Fact]
    public async Task RegistrarIngreso_DatosValidos_InsertaYRetornaRecibo()
    {
        var repo = new Mock<IReciboIngresoRepository>();
        repo.Setup(r => r.ObtenerMaximoCorrelativoAnioAsync(It.IsAny<string>(), default))
            .ReturnsAsync((string?)null);
        repo.Setup(r => r.InsertarAsync(It.IsAny<ReciboIngreso>(), default))
            .Returns(Task.CompletedTask);

        var handler = new RegistrarIngresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarIngresoCommand(
            DateTime.Today, 150m, "Ingreso efectivo",
            Turno: "TUR001", Caja: "01", Usuario: "ADMIN"));

        Assert.True(result.EsExitoso);
        Assert.Equal(150m, result.Valor.Monto);
        Assert.Equal("01", result.Valor.Estado);
        repo.Verify(r => r.InsertarAsync(It.IsAny<ReciboIngreso>(), default), Times.Once);
    }

    [Fact]
    public async Task RegistrarIngreso_MontoInvalido_RetornaFallo()
    {
        // BR-RECIBO-011
        var repo    = new Mock<IReciboIngresoRepository>();
        repo.Setup(r => r.ObtenerMaximoCorrelativoAnioAsync(It.IsAny<string>(), default))
            .ReturnsAsync((string?)null);

        var handler = new RegistrarIngresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarIngresoCommand(
            DateTime.Today, 0m, "Desc"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RECIBO_MONTO_INVALIDO", result.CodigoError);
    }

    [Fact]
    public async Task RegistrarIngreso_SinCorrelativoPrevio_GeneraPrimerCorrelativo()
    {
        // BR-RECIBO-001
        var repo = new Mock<IReciboIngresoRepository>();
        repo.Setup(r => r.ObtenerMaximoCorrelativoAnioAsync(It.IsAny<string>(), default))
            .ReturnsAsync((string?)null);
        repo.Setup(r => r.InsertarAsync(It.IsAny<ReciboIngreso>(), default)).Returns(Task.CompletedTask);

        var anoCorto = DateTime.Now.ToString("yy");
        var handler  = new RegistrarIngresoHandler(repo.Object);
        var result   = await handler.HandleAsync(new RegistrarIngresoCommand(DateTime.Today, 100m, "Desc"));

        Assert.True(result.EsExitoso);
        Assert.StartsWith(anoCorto, result.Valor.CodigoRecibo);
        Assert.Equal(anoCorto + "00000001", result.Valor.CodigoRecibo);
    }

    [Fact]
    public async Task RegistrarIngreso_ConCorrelativoPrevio_IncrementaSecuencia()
    {
        // BR-RECIBO-001: correlativo siguiente
        var anoCorto = DateTime.Now.ToString("yy");
        var repo = new Mock<IReciboIngresoRepository>();
        repo.Setup(r => r.ObtenerMaximoCorrelativoAnioAsync(anoCorto, default))
            .ReturnsAsync(anoCorto + "00000005");
        repo.Setup(r => r.InsertarAsync(It.IsAny<ReciboIngreso>(), default)).Returns(Task.CompletedTask);

        var handler = new RegistrarIngresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarIngresoCommand(DateTime.Today, 100m, "Desc"));

        Assert.True(result.EsExitoso);
        Assert.Equal(anoCorto + "00000006", result.Valor.CodigoRecibo);
    }

    // ── AnularIngresoHandler ─────────────────────────────────────────────────

    [Fact]
    public async Task AnularIngreso_ReciboExistente_AnulaYActualiza()
    {
        var recibo = ReciboIngreso.Registrar("2600000001", DateTime.Today, 100m, "Test");

        var repo = new Mock<IReciboIngresoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("2600000001", default)).ReturnsAsync(recibo);
        repo.Setup(r => r.ActualizarEstadoAsync("2600000001", "04", default)).Returns(Task.CompletedTask);

        var handler = new AnularIngresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new AnularIngresoCommand("2600000001"));

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.ActualizarEstadoAsync("2600000001", "04", default), Times.Once);
    }

    [Fact]
    public async Task AnularIngreso_ReciboNoEncontrado_RetornaFallo()
    {
        var repo = new Mock<IReciboIngresoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync(It.IsAny<string>(), default))
            .ReturnsAsync((ReciboIngreso?)null);

        var handler = new AnularIngresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new AnularIngresoCommand("INEXISTENTE"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RECIBO_INGRESO_NO_ENCONTRADO", result.CodigoError);
    }
}
