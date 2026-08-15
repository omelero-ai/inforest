using Inforest.Application.Caja;
using Inforest.Domain.Entities.Hardware;
using Moq;

namespace Inforest.Application.Tests.Caja;

public sealed class ObtenerTerminalesPinPadHandlerTests
{
    [Fact]
    public async Task HandleAsync_CajaValida_RetornaTerminales()
    {
        var repository = new Mock<IPinPadTerminalRepository>();
        repository.Setup(r => r.ObtenerActivosPorCajaAsync("CAJ01", It.IsAny<CancellationToken>()))
            .ReturnsAsync([
                new TerminalPinPad("TERM-01", "Terminal principal", "01"),
                new TerminalPinPad("TERM-02", "Terminal backup", "01")
            ]);

        var handler = new ObtenerTerminalesPinPadHandler(repository.Object);
        var result = await handler.HandleAsync(new ObtenerTerminalesPinPadQuery("CAJ01"));

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count);
    }

    [Fact]
    public async Task HandleAsync_CajaVacia_RetornaError()
    {
        var repository = new Mock<IPinPadTerminalRepository>();
        var handler = new ObtenerTerminalesPinPadHandler(repository.Object);

        var result = await handler.HandleAsync(new ObtenerTerminalesPinPadQuery(string.Empty));

        Assert.False(result.EsExitoso);
        Assert.Equal("PINPAD_CAJA_REQUERIDA", result.CodigoError);
    }
}
