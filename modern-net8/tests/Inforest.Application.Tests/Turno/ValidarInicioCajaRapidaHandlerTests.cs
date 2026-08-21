using Inforest.Application.Turno;
using Moq;

namespace Inforest.Application.Tests.Turno;

public sealed class ValidarInicioCajaRapidaHandlerTests
{
    [Fact]
    public async Task SinTipoCambioDelDia_NoPermiteIngresoDirecto()
    {
        var turnoRepo = new Mock<ITurnoRepository>(MockBehavior.Strict);
        var tipoCambioRepo = new Mock<ITipoCambioRepository>();
        tipoCambioRepo.Setup(x => x.ObtenerDelDiaAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync((TipoCambioDelDia?)null);

        var handler = new ValidarInicioCajaRapidaHandler(turnoRepo.Object, tipoCambioRepo.Object);

        var result = await handler.HandleAsync(new ValidarInicioCajaRapidaQuery("CA1", "USR1"));

        Assert.True(result.EsExitoso);
        Assert.False(result.Valor!.PermiteIngresoDirecto);
        turnoRepo.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task TurnoCerrado_NoPermiteIngresoDirecto()
    {
        var turnoRepo = new Mock<ITurnoRepository>();
        turnoRepo.Setup(x => x.ObtenerUltimoTurnoAsync("CA1", "USR1", ModoConsultaTurno.PorCajaYUsuario, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new TurnoExistente("2600000001", "CA1", "USR1", new DateTime(2026, 8, 21), true, 100m, 25m));

        var tipoCambioRepo = new Mock<ITipoCambioRepository>();
        tipoCambioRepo.Setup(x => x.ObtenerDelDiaAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new TipoCambioDelDia(3.75m, 3.70m, 0m, 0m));

        var handler = new ValidarInicioCajaRapidaHandler(turnoRepo.Object, tipoCambioRepo.Object);

        var result = await handler.HandleAsync(new ValidarInicioCajaRapidaQuery("CA1", "USR1", ModoConsultaTurno.PorCajaYUsuario));

        Assert.True(result.EsExitoso);
        Assert.False(result.Valor!.PermiteIngresoDirecto);
        turnoRepo.Verify(x => x.ReAperturarAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<decimal>(), It.IsAny<decimal>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task TurnoAbiertoYTipoCambio_ReutilizaTurnoYPermiteIngresoDirecto()
    {
        var turnoRepo = new Mock<ITurnoRepository>();
        turnoRepo.Setup(x => x.ObtenerUltimoTurnoAsync("CA1", "USR1", ModoConsultaTurno.PorCajaYUsuario, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new TurnoExistente("2600000009", "CA1", "USR0", new DateTime(2026, 8, 21), false, 150m, 20m));
        turnoRepo.Setup(x => x.ReAperturarAsync("2600000009", "USR1", 150m, 20m, It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var tipoCambioRepo = new Mock<ITipoCambioRepository>();
        tipoCambioRepo.Setup(x => x.ObtenerDelDiaAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new TipoCambioDelDia(3.75m, 3.70m, 0m, 0m));

        var handler = new ValidarInicioCajaRapidaHandler(turnoRepo.Object, tipoCambioRepo.Object);

        var result = await handler.HandleAsync(new ValidarInicioCajaRapidaQuery("CA1", "USR1", ModoConsultaTurno.PorCajaYUsuario));

        Assert.True(result.EsExitoso);
        Assert.True(result.Valor!.PermiteIngresoDirecto);
        Assert.Equal("2600000009", result.Valor.CodigoTurno);
        turnoRepo.VerifyAll();
    }
}
