using Inforest.Application.Motorizado;
using Inforest.Domain.Entities.Motorizado;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Motorizado;

/// <summary>
/// Tests de handlers de Motorizado.
/// Legacy: frmAsignacionMotorizado.frm, frmLlegadaSalida.frm, frmTarifaMotorizado.frm
/// Reglas: BR-DEL-001, BR-DEL-002, BR-DEL-005, BR-DEL-006.
/// </summary>
public class MotorizadoHandlersTests
{
    [Fact]
    public async Task AsignarMotorizado_Todo_OK_RetornaOk()
    {
        var motorizado = Inforest.Domain.Entities.Motorizado.Motorizado.Crear("MOT001", "12345678", 10m, 15m, 20m);
        var pedido = Inforest.Domain.Entities.Delivery.PedidoDelivery.Crear(
            "PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);

        var mockMotRepo = new Mock<IMotorizadoRepository>();
        mockMotRepo.Setup(r => r.ObtenerPorCodigoAsync("MOT001", It.IsAny<CancellationToken>()))
                   .ReturnsAsync(motorizado);
        mockMotRepo.Setup(r => r.RegistrarAsignacionAsync(It.IsAny<AsignacionMotorizado>(), It.IsAny<CancellationToken>()))
                   .Returns(Task.CompletedTask);

        var mockPedRepo = new Mock<IPedidoDeliveryRepository>();
        mockPedRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
                   .ReturnsAsync(pedido);
        mockPedRepo.Setup(r => r.AsignarMotorizadoAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
                   .Returns(Task.CompletedTask);

        var handler = new AsignarMotorizadoHandler(mockMotRepo.Object, mockPedRepo.Object);
        var result = await handler.HandleAsync(new AsignarMotorizadoCommand("PED001", "MOT001", "USR001"));

        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task AsignarMotorizado_MotorizadoNoExiste_RetornaFallo()
    {
        var mockMotRepo = new Mock<IMotorizadoRepository>();
        mockMotRepo.Setup(r => r.ObtenerPorCodigoAsync("MOTXXX", It.IsAny<CancellationToken>()))
                   .ReturnsAsync((Inforest.Domain.Entities.Motorizado.Motorizado?)null);

        var mockPedRepo = new Mock<IPedidoDeliveryRepository>();

        var handler = new AsignarMotorizadoHandler(mockMotRepo.Object, mockPedRepo.Object);
        var result = await handler.HandleAsync(new AsignarMotorizadoCommand("PED001", "MOTXXX", "USR001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("MOTORIZADO_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerTarifa_TipoDiaLV_RetornaTarifaCorrectamente()
    {
        var motorizado = Inforest.Domain.Entities.Motorizado.Motorizado.Crear("MOT001", "12345678", 12m, 18m, 25m);

        var mockRepo = new Mock<IMotorizadoRepository>();
        mockRepo.Setup(r => r.ObtenerPorCodigoAsync("MOT001", It.IsAny<CancellationToken>()))
                .ReturnsAsync(motorizado);

        var handler = new ObtenerTarifaMotorizadoHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new ObtenerTarifaMotorizadoQuery("MOT001", TipoDiaTarifa.LunesViernes));

        Assert.True(result.EsExitoso);
        Assert.Equal(12m, result.Valor);
    }

    [Fact]
    public async Task RegistrarLlegada_LlamaaRepo()
    {
        var mockRepo = new Mock<IMotorizadoRepository>();
        mockRepo.Setup(r => r.RegistrarLlegadaAsync("PED001", "MOT001", It.IsAny<CancellationToken>()))
                .Returns(Task.CompletedTask);

        var handler = new RegistrarLlegadaSalidaHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new RegistrarLlegadaSalidaCommand("PED001", "MOT001", EsLlegada: true));

        Assert.True(result.EsExitoso);
        mockRepo.Verify(r => r.RegistrarLlegadaAsync("PED001", "MOT001", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task RegistrarSalida_LlamaaRepo()
    {
        var mockRepo = new Mock<IMotorizadoRepository>();
        mockRepo.Setup(r => r.RegistrarSalidaAsync("PED001", "MOT001", It.IsAny<CancellationToken>()))
                .Returns(Task.CompletedTask);

        var handler = new RegistrarLlegadaSalidaHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new RegistrarLlegadaSalidaCommand("PED001", "MOT001", EsLlegada: false));

        Assert.True(result.EsExitoso);
        mockRepo.Verify(r => r.RegistrarSalidaAsync("PED001", "MOT001", It.IsAny<CancellationToken>()), Times.Once);
    }
}
