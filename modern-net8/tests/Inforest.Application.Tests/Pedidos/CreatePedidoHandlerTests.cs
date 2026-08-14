using Inforest.Application.Pedidos;
using Inforest.Application.Turno;
using Inforest.Domain.Entities.Configuracion;
using TurnoEntity = Inforest.Domain.Entities.Configuracion.Turno;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Pedidos;

public class CreatePedidoHandlerTests
{
    [Fact]
    public async Task ComandoValido_LlamaRepositoryInsertarAsync()
    {
        var pedidoRepository = new Mock<IPedidoRepository>();
        var turnoRepository = new Mock<ITurnoRepository>();
        turnoRepository
            .Setup(x => x.ObtenerTurnoActualAsync("CAJ1", It.IsAny<CancellationToken>()))
            .ReturnsAsync(TurnoEntity.Abrir("2408120001", "CAJ1", "USR1", new DateTime(2026, 8, 12), 50m));

        var handler = new CreatePedidoHandler(pedidoRepository.Object, turnoRepository.Object);
        var command = new CreatePedidoCommand(
            "CAJ1",
            "2408120001",
            "M01",
            "MOZ1",
            "USR1",
            CanalVenta.Local,
            2,
            0,
            "Observación",
            Array.Empty<CreateDetalleItem>());

        var result = await handler.HandleAsync(command);

        Assert.True(result.EsExitoso);
        pedidoRepository.Verify(x => x.InsertarAsync(It.IsAny<Pedido>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task TurnoVacio_RetornaError()
    {
        var pedidoRepository = new Mock<IPedidoRepository>();
        var turnoRepository = new Mock<ITurnoRepository>();
        var handler = new CreatePedidoHandler(pedidoRepository.Object, turnoRepository.Object);

        var command = new CreatePedidoCommand(
            "CAJ1",
            string.Empty,
            "M01",
            "MOZ1",
            "USR1",
            CanalVenta.Local,
            1,
            0,
            null,
            Array.Empty<CreateDetalleItem>());

        var result = await handler.HandleAsync(command);

        Assert.False(result.EsExitoso);
        Assert.Equal("PEDIDO_TURNO_REQUERIDO", result.CodigoError);
        pedidoRepository.Verify(x => x.InsertarAsync(It.IsAny<Pedido>(), It.IsAny<CancellationToken>()), Times.Never);
    }
}
