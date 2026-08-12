using Inforest.Application.Delivery;
using Inforest.Application.Motorizado;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Entities.Motorizado;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Delivery;

/// <summary>
/// Tests de handlers de Delivery.
/// Legacy: frmNuevoDelivery.frm, frmDespachador.frm, frmAsignacionMotorizado.frm
/// </summary>
public class DeliveryHandlersTests
{
    // ── CrearClienteDeliveryHandler ────────────────────────────────────────────

    [Fact]
    public async Task CrearCliente_ClienteNuevo_RetornaOk()
    {
        var mockRepo = new Mock<IClienteDeliveryRepository>();
        mockRepo.Setup(r => r.ObtenerPorCodigoAsync("DEL001", It.IsAny<CancellationToken>()))
                .ReturnsAsync((ClienteDelivery?)null);
        mockRepo.Setup(r => r.InsertarAsync(It.IsAny<ClienteDelivery>(), It.IsAny<CancellationToken>()))
                .Returns(Task.CompletedTask);

        var handler = new CrearClienteDeliveryHandler(mockRepo.Object);
        var cmd = new CrearClienteDeliveryCommand(
            "DEL001", "01", "García", "Juan", "987654321",
            "Av Lima", null, null, null, null, null, null, null, null, null);

        var result = await handler.HandleAsync(cmd);

        Assert.True(result.EsExitoso);
        mockRepo.Verify(r => r.InsertarAsync(It.IsAny<ClienteDelivery>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task CrearCliente_ClienteExistente_RetornaFallo()
    {
        var clienteExistente = ClienteDelivery.Crear("DEL001", null, "García", null, "987654321", null, null, null);
        var mockRepo = new Mock<IClienteDeliveryRepository>();
        mockRepo.Setup(r => r.ObtenerPorCodigoAsync("DEL001", It.IsAny<CancellationToken>()))
                .ReturnsAsync(clienteExistente);

        var handler = new CrearClienteDeliveryHandler(mockRepo.Object);
        var cmd = new CrearClienteDeliveryCommand(
            "DEL001", null, "García", null, "987654321",
            null, null, null, null, null, null, null, null, null, null);

        var result = await handler.HandleAsync(cmd);

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_CLIENTE_YA_EXISTE", result.CodigoError);
    }

    // ── ActualizarEstadoDeliveryHandler ────────────────────────────────────────

    [Fact]
    public async Task ActualizarEstado_PedidoExistente_CentralDisponible_RetornaOk()
    {
        var pedido = PedidoDelivery.Crear("PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);

        var mockPedidoRepo = new Mock<IPedidoDeliveryRepository>();
        mockPedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
                      .ReturnsAsync(pedido);
        mockPedidoRepo.Setup(r => r.ActualizarEstadoAsync(It.IsAny<string>(), It.IsAny<EstadoDelivery>(), It.IsAny<CancellationToken>()))
                      .Returns(Task.CompletedTask);

        var mockCentral = new Mock<ICentralPedidosRepository>();
        mockCentral.Setup(r => r.VerificarConexionAsync(It.IsAny<CancellationToken>()))
                   .ReturnsAsync(true);
        mockCentral.Setup(r => r.ModificarEstadoDeliveryAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
                   .Returns(Task.CompletedTask);

        var handler = new ActualizarEstadoDeliveryHandler(mockPedidoRepo.Object, mockCentral.Object);
        var cmd = new ActualizarEstadoDeliveryCommand("PED001", EstadoDelivery.EnPreparacion, "USR001");

        var result = await handler.HandleAsync(cmd);

        Assert.True(result.EsExitoso);
        mockCentral.Verify(r => r.ModificarEstadoDeliveryAsync("PED001", "EnPreparacion", "USR001", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ActualizarEstado_PedidoInexistente_RetornaFallo()
    {
        var mockPedidoRepo = new Mock<IPedidoDeliveryRepository>();
        mockPedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED999", It.IsAny<CancellationToken>()))
                      .ReturnsAsync((PedidoDelivery?)null);

        var mockCentral = new Mock<ICentralPedidosRepository>();

        var handler = new ActualizarEstadoDeliveryHandler(mockPedidoRepo.Object, mockCentral.Object);
        var result = await handler.HandleAsync(new ActualizarEstadoDeliveryCommand("PED999", EstadoDelivery.EnPreparacion, "USR001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_PEDIDO_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task ActualizarEstado_CentralNoDisponible_NoLlamaCentral()
    {
        var pedido = PedidoDelivery.Crear("PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);

        var mockPedidoRepo = new Mock<IPedidoDeliveryRepository>();
        mockPedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
                      .ReturnsAsync(pedido);
        mockPedidoRepo.Setup(r => r.ActualizarEstadoAsync(It.IsAny<string>(), It.IsAny<EstadoDelivery>(), It.IsAny<CancellationToken>()))
                      .Returns(Task.CompletedTask);

        var mockCentral = new Mock<ICentralPedidosRepository>();
        mockCentral.Setup(r => r.VerificarConexionAsync(It.IsAny<CancellationToken>()))
                   .ReturnsAsync(false);

        var handler = new ActualizarEstadoDeliveryHandler(mockPedidoRepo.Object, mockCentral.Object);
        var result = await handler.HandleAsync(new ActualizarEstadoDeliveryCommand("PED001", EstadoDelivery.EnPreparacion, "USR001"));

        Assert.True(result.EsExitoso);
        mockCentral.Verify(r => r.ModificarEstadoDeliveryAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()), Times.Never);
    }
}
