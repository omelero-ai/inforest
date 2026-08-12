using Inforest.Application.Caja;
using Moq;

namespace Inforest.Application.Tests.Caja;

public class ObtenerResumenTurnoHandlerTests
{
    [Fact]
    public async Task HandleAsync_DevuelveResumenDelTurno()
    {
        var resumenEsperado = new ResumenTurnoCaja("TUR001", 500m, 480m, 20m, 5);
        var pagoRepo = new Mock<IPagoRepository>();
        pagoRepo.Setup(r => r.ObtenerResumenTurnoAsync("TUR001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(resumenEsperado);

        var handler = new ObtenerResumenTurnoHandler(pagoRepo.Object);
        var result = await handler.HandleAsync(new ObtenerResumenTurnoQuery("TUR001"));

        Assert.True(result.EsExitoso);
        Assert.NotNull(result.Valor);
        Assert.Equal("TUR001", result.Valor!.CodigoTurno);
        Assert.Equal(500m, result.Valor.TotalVentas);
        Assert.Equal(5, result.Valor.DocumentosCobrados);
        pagoRepo.Verify(r => r.ObtenerResumenTurnoAsync("TUR001", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task HandleAsync_ConTurnoSinVentas_DevuelveResumenVacio()
    {
        var resumenVacio = new ResumenTurnoCaja("TUR002", 0m, 0m, 0m, 0);
        var pagoRepo = new Mock<IPagoRepository>();
        pagoRepo.Setup(r => r.ObtenerResumenTurnoAsync("TUR002", It.IsAny<CancellationToken>()))
            .ReturnsAsync(resumenVacio);

        var handler = new ObtenerResumenTurnoHandler(pagoRepo.Object);
        var result = await handler.HandleAsync(new ObtenerResumenTurnoQuery("TUR002"));

        Assert.True(result.EsExitoso);
        Assert.Equal(0, result.Valor!.DocumentosCobrados);
        Assert.Equal(0m, result.Valor.TotalVentas);
    }
}
