using Inforest.Application.Pedidos;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Pedidos;

/// <summary>
/// Tests del handler UpdatePedido.
/// Legacy: spUpd_MPEDIDO, frmPedido.frm — modificación de estado/observación de pedido.
/// Reglas: BR-PEDIDO-001, BR-PEDIDO-002, BR-SQL-002
/// </summary>
public class UpdatePedidoHandlerTests
{
    private static Pedido PedidoAbierto() =>
        Pedido.Crear("PED001", "M01", "TURNO01", "USR01", CanalVenta.Local);

    [Fact]
    public async Task CodigoVacio_RetornaError_PEDIDO_CODIGO_REQUERIDO()
    {
        var repo = new Mock<IPedidoRepository>();
        var handler = new UpdatePedidoHandler(repo.Object);

        var result = await handler.HandleAsync(new UpdatePedidoCommand(string.Empty, null, null, "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("PEDIDO_CODIGO_REQUERIDO", result.CodigoError);
        repo.Verify(x => x.ActualizarAsync(It.IsAny<Pedido>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task PedidoNoExiste_RetornaError_PEDIDO_NO_ENCONTRADO()
    {
        var repo = new Mock<IPedidoRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("PED999", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Pedido?)null);

        var handler = new UpdatePedidoHandler(repo.Object);
        var result = await handler.HandleAsync(new UpdatePedidoCommand("PED999", null, null, "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("PEDIDO_NO_ENCONTRADO", result.CodigoError);
        repo.Verify(x => x.ActualizarAsync(It.IsAny<Pedido>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task PedidoExiste_SinCambios_LlamaActualizarAsync()
    {
        var pedido = PedidoAbierto();
        var repo = new Mock<IPedidoRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var handler = new UpdatePedidoHandler(repo.Object);
        var result = await handler.HandleAsync(new UpdatePedidoCommand("PED001", null, null, "USR01"));

        Assert.True(result.EsExitoso);
        repo.Verify(x => x.ActualizarAsync(pedido, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task PedidoExiste_ConObservacion_ActualizaObservacion()
    {
        var pedido = PedidoAbierto();
        var repo = new Mock<IPedidoRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var handler = new UpdatePedidoHandler(repo.Object);
        var result = await handler.HandleAsync(new UpdatePedidoCommand("PED001", null, "Sin sal por favor", "USR01"));

        Assert.True(result.EsExitoso);
        Assert.Equal("Sin sal por favor", pedido.Observacion);
        repo.Verify(x => x.ActualizarAsync(pedido, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task PedidoExiste_ConNuevoEstadoCerrado_CambiaEstado()
    {
        var pedido = PedidoAbierto();
        var repo = new Mock<IPedidoRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var handler = new UpdatePedidoHandler(repo.Object);
        // Legacy: "CE" = Cerrado, "03" también es Cerrado (ver LegacyMappings.MapEstado)
        var result = await handler.HandleAsync(new UpdatePedidoCommand("PED001", "CE", null, "USR01"));

        Assert.True(result.EsExitoso);
        Assert.Equal(EstadoPedido.Cerrado, pedido.Estado);
        repo.Verify(x => x.ActualizarAsync(pedido, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task PedidoExiste_ConNuevoEstadoAnulado_CambiaEstado()
    {
        var pedido = PedidoAbierto();
        var repo = new Mock<IPedidoRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var handler = new UpdatePedidoHandler(repo.Object);
        // Legacy: "AN" = Anulado, "04" también es Anulado (ver LegacyMappings.MapEstado)
        var result = await handler.HandleAsync(new UpdatePedidoCommand("PED001", "AN", null, "USR01"));

        Assert.True(result.EsExitoso);
        Assert.Equal(EstadoPedido.Anulado, pedido.Estado);
    }

    [Fact]
    public async Task PedidoExiste_EstadoNull_MantieneEstadoOriginal()
    {
        var pedido = PedidoAbierto();
        var estadoOriginal = pedido.Estado;
        var repo = new Mock<IPedidoRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var handler = new UpdatePedidoHandler(repo.Object);
        await handler.HandleAsync(new UpdatePedidoCommand("PED001", null, null, "USR01"));

        Assert.Equal(estadoOriginal, pedido.Estado);
    }
}
