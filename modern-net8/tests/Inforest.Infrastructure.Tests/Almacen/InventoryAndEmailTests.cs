using Inforest.Application.Interfaces;
using Moq;

namespace Inforest.Infrastructure.Tests.Almacen;

/// <summary>
/// Tests de contrato del gateway de almacén usando Moq.
/// Legacy: clsAlmacen.cls — BR-008.
/// </summary>
public sealed class InventoryGatewayTests
{
    [Fact]
    public async Task EjecutarDescargoVentaAsync_RetornaTrue_CuandoSpEjecutaOk()
    {
        var mock = new Mock<IInventoryGateway>();
        mock.Setup(g => g.EjecutarDescargoVentaAsync(
                It.IsAny<string>(), It.IsAny<DateTime>(), It.IsAny<DateTime>(),
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<int>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var resultado = await mock.Object.EjecutarDescargoVentaAsync(
            "ALM01", DateTime.Today, DateTime.Today, "#TEMP001", "LOC01", "PED001", 1);

        Assert.True(resultado);
    }

    [Fact]
    public async Task EjecutarDescargoVentaPorInsumoAsync_RetornaTrue_CuandoSpEjecutaOk()
    {
        var mock = new Mock<IInventoryGateway>();
        mock.Setup(g => g.EjecutarDescargoVentaPorInsumoAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<DateTime>(),
                It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var resultado = await mock.Object.EjecutarDescargoVentaPorInsumoAsync(
            "PED001", "ALM01", DateTime.Today, "USR01");

        Assert.True(resultado);
    }

    [Fact]
    public async Task ModificarInsumosAsync_RetornaTrue_CuandoSpEjecutaOk()
    {
        var mock = new Mock<IInventoryGateway>();
        mock.Setup(g => g.ModificarInsumosAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var resultado = await mock.Object.ModificarInsumosAsync("PED001", "ALM01", "USR01");

        Assert.True(resultado);
    }

    [Fact]
    public async Task InsertarLogDescargoAsync_SeInvocaConParametrosCorrectosLegacy()
    {
        var mock = new Mock<IInventoryGateway>();
        mock.Setup(g => g.InsertarLogDescargoAsync(
                It.IsAny<int>(), It.IsAny<DateTime>(), It.IsAny<string>(),
                It.IsAny<DateTime>(), It.IsAny<DateTime>(),
                It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>(),
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        await mock.Object.InsertarLogDescargoAsync(
            1, DateTime.Now, "USR01",
            DateTime.Today, DateTime.Today,
            10, 9, 1,
            "192.168.1.1", "EQUIPO01", "user");

        mock.Verify(g => g.InsertarLogDescargoAsync(
            1, It.IsAny<DateTime>(), "USR01",
            It.IsAny<DateTime>(), It.IsAny<DateTime>(),
            10, 9, 1,
            "192.168.1.1", "EQUIPO01", "user",
            It.IsAny<CancellationToken>()), Times.Once);
    }
}

/// <summary>
/// Tests de contrato del servicio de email usando Moq.
/// Legacy: claCorreoElectronico.cls.
/// </summary>
public sealed class NotificacionEmailServiceTests
{
    [Fact]
    public async Task EnviarCorreoAsync_RetornaTrue_CuandoEnvioExitoso()
    {
        var mock = new Mock<INotificacionEmailService>();
        mock.Setup(s => s.EnviarCorreoAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var resultado = await mock.Object.EnviarCorreoAsync(
            "destinatario@test.com", "Asunto", "Cuerpo");

        Assert.True(resultado);
    }

    [Fact]
    public async Task EnviarCorreoProrrogaAsync_SeInvocaSinExcepcion()
    {
        var mock = new Mock<INotificacionEmailService>();
        mock.Setup(s => s.EnviarCorreoProrrogaAsync(
                It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        await mock.Object.EnviarCorreoProrrogaAsync("admin@test.com");

        mock.Verify(s => s.EnviarCorreoProrrogaAsync("admin@test.com", It.IsAny<CancellationToken>()), Times.Once);
    }
}
