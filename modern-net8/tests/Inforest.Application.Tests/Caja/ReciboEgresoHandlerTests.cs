using Inforest.Application.Caja;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Caja;

/// <summary>
/// Tests de handlers para recibos de egreso.
/// Cubre BR-RECIBO-007, BR-RECIBO-008, BR-RECIBO-010, BR-RECIBO-011, BR-RECIBO-012.
/// </summary>
public class ReciboEgresoHandlerTests
{
    // ── ObtenerEgresosHandler ────────────────────────────────────────────────

    [Fact]
    public async Task ObtenerEgresos_RangoValido_RetornaLista()
    {
        var repo = new Mock<IReciboEgresoRepository>();
        var lista = new List<ReciboEgreso>
        {
            ReciboEgreso.Registrar("2600000001", DateTime.Today, 50m, "Pago luz", tipoEgreso: "SRV")
        };
        repo.Setup(r => r.ObtenerPorFechaAsync(It.IsAny<DateTime>(), It.IsAny<DateTime>(), null, default))
            .ReturnsAsync(lista.AsReadOnly());

        var handler = new ObtenerEgresosHandler(repo.Object);
        var result  = await handler.HandleAsync(new ObtenerEgresosQuery(DateTime.Today.AddDays(-1), DateTime.Today));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor);
    }

    [Fact]
    public async Task ObtenerEgresos_FechaDesdeEsMayor_RetornaFallo()
    {
        var repo    = new Mock<IReciboEgresoRepository>();
        var handler = new ObtenerEgresosHandler(repo.Object);
        var result  = await handler.HandleAsync(
            new ObtenerEgresosQuery(DateTime.Today.AddDays(1), DateTime.Today));

        Assert.False(result.EsExitoso);
        Assert.Equal("EGRESO_RANGO_FECHAS_INVALIDO", result.CodigoError);
    }

    // ── RegistrarEgresoHandler ───────────────────────────────────────────────

    [Fact]
    public async Task RegistrarEgreso_DatosValidos_InsertaYRetornaEgreso()
    {
        var repo = new Mock<IReciboEgresoRepository>();
        repo.Setup(r => r.ObtenerMaximoCorrelativoAnioAsync(It.IsAny<string>(), default))
            .ReturnsAsync((string?)null);
        repo.Setup(r => r.InsertarAsync(It.IsAny<ReciboEgreso>(), default)).Returns(Task.CompletedTask);

        var handler = new RegistrarEgresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarEgresoCommand(
            DateTime.Today, 80m, "Pago proveedor",
            ModoEgreso: "VARIOS", TipoEgreso: "GEN",
            Turno: "TUR001", Caja: "01", Usuario: "ADMIN"));

        Assert.True(result.EsExitoso);
        Assert.Equal(80m, result.Valor.Monto);
        Assert.Equal("VARIOS", result.Valor.ModoEgreso);
        Assert.Equal("01", result.Valor.Estado);
        repo.Verify(r => r.InsertarAsync(It.IsAny<ReciboEgreso>(), default), Times.Once);
    }

    [Fact]
    public async Task RegistrarEgreso_MontoInvalido_RetornaFallo()
    {
        // BR-RECIBO-011
        var repo = new Mock<IReciboEgresoRepository>();
        repo.Setup(r => r.ObtenerMaximoCorrelativoAnioAsync(It.IsAny<string>(), default))
            .ReturnsAsync((string?)null);

        var handler = new RegistrarEgresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarEgresoCommand(DateTime.Today, 0m, "Desc"));

        Assert.False(result.EsExitoso);
        Assert.Equal("EGRESO_MONTO_INVALIDO", result.CodigoError);
    }

    [Fact]
    public async Task RegistrarEgreso_ModoNCsinReferencia_RetornaFallo()
    {
        // BR-RECIBO-012
        var repo = new Mock<IReciboEgresoRepository>();
        repo.Setup(r => r.ObtenerMaximoCorrelativoAnioAsync(It.IsAny<string>(), default))
            .ReturnsAsync((string?)null);

        var handler = new RegistrarEgresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarEgresoCommand(
            DateTime.Today, 50m, "NC sin ref", ModoEgreso: "NOTACREDITO"));

        Assert.False(result.EsExitoso);
        Assert.Equal("EGRESO_REFERENCIA_NC_REQUERIDA", result.CodigoError);
    }

    [Fact]
    public async Task RegistrarEgreso_CorrelatiivoSecuencial_Incrementa()
    {
        // BR-RECIBO-007: correlativo YY + 8 dígitos
        var anoCorto = DateTime.Now.ToString("yy");
        var repo = new Mock<IReciboEgresoRepository>();
        repo.Setup(r => r.ObtenerMaximoCorrelativoAnioAsync(anoCorto, default))
            .ReturnsAsync(anoCorto + "00000003");
        repo.Setup(r => r.InsertarAsync(It.IsAny<ReciboEgreso>(), default)).Returns(Task.CompletedTask);

        var handler = new RegistrarEgresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarEgresoCommand(DateTime.Today, 80m, "Desc"));

        Assert.True(result.EsExitoso);
        Assert.Equal(anoCorto + "00000004", result.Valor.CodigoRecibo);
    }

    // ── AnularEgresoHandler ──────────────────────────────────────────────────

    [Fact]
    public async Task AnularEgreso_ReciboExistente_AnulaYActualiza()
    {
        var recibo = ReciboEgreso.Registrar("2600000001", DateTime.Today, 80m, "Test");

        var repo = new Mock<IReciboEgresoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("2600000001", default)).ReturnsAsync(recibo);
        repo.Setup(r => r.ActualizarEstadoAsync("2600000001", "04", default)).Returns(Task.CompletedTask);

        var handler = new AnularEgresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new AnularEgresoCommand("2600000001"));

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.ActualizarEstadoAsync("2600000001", "04", default), Times.Once);
    }

    [Fact]
    public async Task AnularEgreso_ReciboNoEncontrado_RetornaFallo()
    {
        var repo = new Mock<IReciboEgresoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync(It.IsAny<string>(), default))
            .ReturnsAsync((ReciboEgreso?)null);

        var handler = new AnularEgresoHandler(repo.Object);
        var result  = await handler.HandleAsync(new AnularEgresoCommand("INEXISTENTE"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RECIBO_EGRESO_NO_ENCONTRADO", result.CodigoError);
    }
}
