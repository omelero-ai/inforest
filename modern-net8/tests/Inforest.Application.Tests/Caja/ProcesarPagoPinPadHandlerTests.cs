using Inforest.Application.Caja;
using Inforest.Application.Interfaces.Hardware;
using Moq;

namespace Inforest.Application.Tests.Caja;

public sealed class ProcesarPagoPinPadHandlerTests
{
    [Fact]
    public async Task HandleAsync_VentaExitosa_RetornaReferencia()
    {
        var service = new Mock<IPinPadService>();
        service.Setup(s => s.InicializarAsync(null, It.IsAny<CancellationToken>()))
            .ReturnsAsync(PinPadResult.Ok("Init"));
        service.Setup(s => s.EjecutarVentaAsync(25m, 1, 60, It.IsAny<CancellationToken>()))
            .ReturnsAsync(PinPadResult.Ok("Venta OK", "AUTH-001", "OP-001"));
        service.Setup(s => s.CerrarPuertoAsync(It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var handler = new ProcesarPagoPinPadHandler(service.Object);
        var result = await handler.HandleAsync(new ProcesarPagoPinPadCommand(25m, 1, "TERM-01"));

        Assert.True(result.EsExitoso);
        Assert.Equal("OP-001", result.Valor!.Referencia);
        Assert.Equal("TERM-01", result.Valor.CodigoTerminal);
        service.Verify(s => s.CerrarPuertoAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task HandleAsync_MontoInvalido_RetornaError()
    {
        var service = new Mock<IPinPadService>();
        var handler = new ProcesarPagoPinPadHandler(service.Object);

        var result = await handler.HandleAsync(new ProcesarPagoPinPadCommand(0m));

        Assert.False(result.EsExitoso);
        Assert.Equal("PINPAD_MONTO_INVALIDO", result.CodigoError);
    }

    [Fact]
    public async Task HandleAsync_InicializacionFallida_NoEjecutaVenta()
    {
        var service = new Mock<IPinPadService>();
        service.Setup(s => s.InicializarAsync(null, It.IsAny<CancellationToken>()))
            .ReturnsAsync(PinPadResult.ErrorComunicacion(-2, "Sin puerto"));

        var handler = new ProcesarPagoPinPadHandler(service.Object);
        var result = await handler.HandleAsync(new ProcesarPagoPinPadCommand(10m));

        Assert.False(result.EsExitoso);
        Assert.Equal("PINPAD_INICIALIZACION_FALLIDA", result.CodigoError);
        service.Verify(s => s.EjecutarVentaAsync(It.IsAny<decimal>(), It.IsAny<int>(), It.IsAny<int>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task HandleAsync_OperacionRechazada_RetornaError()
    {
        var service = new Mock<IPinPadService>();
        service.Setup(s => s.InicializarAsync(null, It.IsAny<CancellationToken>()))
            .ReturnsAsync(PinPadResult.Ok("Init"));
        service.Setup(s => s.EjecutarVentaAsync(10m, 1, 60, It.IsAny<CancellationToken>()))
            .ReturnsAsync(PinPadResult.Rechazado("Tarjeta rechazada"));
        service.Setup(s => s.CerrarPuertoAsync(It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var handler = new ProcesarPagoPinPadHandler(service.Object);
        var result = await handler.HandleAsync(new ProcesarPagoPinPadCommand(10m));

        Assert.False(result.EsExitoso);
        Assert.Equal("PINPAD_OPERACION_RECHAZADA", result.CodigoError);
        service.Verify(s => s.CerrarPuertoAsync(It.IsAny<CancellationToken>()), Times.Once);
    }
}
