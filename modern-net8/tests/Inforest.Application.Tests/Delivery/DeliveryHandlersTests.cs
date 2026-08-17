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
        mockRepo.Setup(r => r.ObtenerPorTelefonoAsync("987654321", It.IsAny<CancellationToken>()))
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
        mockRepo.Setup(r => r.ObtenerPorTelefonoAsync("987654321", It.IsAny<CancellationToken>()))
                .ReturnsAsync((ClienteDelivery?)null);

        var handler = new CrearClienteDeliveryHandler(mockRepo.Object);
        var cmd = new CrearClienteDeliveryCommand(
            "DEL001", null, "García", null, "987654321",
            null, null, null, null, null, null, null, null, null, null);

        var result = await handler.HandleAsync(cmd);

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_CLIENTE_YA_EXISTE", result.CodigoError);
    }

    [Fact]
    public async Task CrearCliente_SinTelefonoNiIdentidad_RetornaFallo()
    {
        var mockRepo = new Mock<IClienteDeliveryRepository>();
        var handler = new CrearClienteDeliveryHandler(mockRepo.Object);
        var cmd = new CrearClienteDeliveryCommand(
            "DEL002", "01", "García", "Juan", null,
            "Av Lima", null, null, null, null, null, null, null, null, null);

        var result = await handler.HandleAsync(cmd);

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_TELEFONO_O_IDENTIDAD_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerSiguienteCodigo_MaximoNumerico_RetornaSiguiente()
    {
        var mockRepo = new Mock<IClienteDeliveryRepository>();
        mockRepo.Setup(r => r.ObtenerMaximoCodigoAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync("0000123");

        var handler = new ObtenerSiguienteCodigoClienteDeliveryHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new ObtenerSiguienteCodigoClienteDeliveryQuery());

        Assert.True(result.EsExitoso);
        Assert.Equal("0000124", result.Valor);
    }

    [Fact]
    public async Task ActualizarFoto_ClienteExistenteYFotoValida_RetornaOk()
    {
        var cliente = ClienteDelivery.Crear("DEL001", null, "García", "Juan", "987654321", "Av Lima", null, null);
        var mockRepo = new Mock<IClienteDeliveryRepository>();
        mockRepo.Setup(r => r.ObtenerPorCodigoAsync("DEL001", It.IsAny<CancellationToken>()))
                .ReturnsAsync(cliente);
        mockRepo.Setup(r => r.ActualizarFotoAsync("DEL001", It.IsAny<byte[]>(), It.IsAny<CancellationToken>()))
                .Returns(Task.CompletedTask);

        var handler = new ActualizarFotoClienteDeliveryHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new ActualizarFotoClienteDeliveryCommand("DEL001", [1, 2, 3]));

        Assert.True(result.EsExitoso);
        mockRepo.Verify(r => r.ActualizarFotoAsync("DEL001", It.IsAny<byte[]>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ActualizarFoto_ClienteInexistente_RetornaFallo()
    {
        var mockRepo = new Mock<IClienteDeliveryRepository>();
        mockRepo.Setup(r => r.ObtenerPorCodigoAsync("DEL404", It.IsAny<CancellationToken>()))
                .ReturnsAsync((ClienteDelivery?)null);

        var handler = new ActualizarFotoClienteDeliveryHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new ActualizarFotoClienteDeliveryCommand("DEL404", [1]));

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_CLIENTE_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerListadoMantenimiento_RetornaItems()
    {
        var mockReadRepo = new Mock<IClienteDeliveryReadRepository>();
        mockReadRepo.Setup(r => r.ListarMantenimientoAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync([
                new ClienteDeliveryListadoItem("0000001", "FRECUENTE", "JUAN PEREZ", "999111222", "AV. 1", 12.5m, 0.5m, "ACTIVO", true)
            ]);

        var handler = new ObtenerClientesDeliveryListadoHandler(mockReadRepo.Object);
        var result = await handler.HandleAsync(new ObtenerClientesDeliveryListadoQuery());

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!);
        Assert.Equal("0000001", result.Valor![0].Codigo);
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

/// <summary>
/// Tests de handlers de Central de Pedidos — frmCentralPedidos.frm.
/// Reglas: BR-DEL-012, BR-DEL-013, BR-DEL-014.
/// </summary>
public class CentralPedidosHandlersTests
{
    // ── ConfirmarEntregaCentralHandler ────────────────────────────────────────

    [Fact]
    public async Task ConfirmarEntrega_PedidoPendiente_SinPago_SinSupervisor_RetornaRequiereSupervisor()
    {
        // BR-DEL-013: pedido con estado POR COBRAR sin autorización → REQUIERE_SUPERVISOR_22
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync(false);
        mock.Setup(r => r.ObtenerEstadoPagoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync("NO PAGADO");

        var handler = new ConfirmarEntregaCentralHandler(mock.Object);
        var result = await handler.HandleAsync(new ConfirmarEntregaCentralCommand("PED001", "USR001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("REQUIERE_SUPERVISOR_22", result.CodigoError);
        mock.Verify(r => r.ConfirmarEntregaAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ConfirmarEntrega_PedidoPendiente_SinPago_ConSupervisor_RetornaOk()
    {
        // BR-DEL-013: con SupervisorAutorizado = true → confirma igual
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync(false);
        mock.Setup(r => r.ObtenerEstadoPagoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync("NO PAGADO");
        mock.Setup(r => r.ConfirmarEntregaAsync("PED001", "USR001", It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var handler = new ConfirmarEntregaCentralHandler(mock.Object);
        var result = await handler.HandleAsync(new ConfirmarEntregaCentralCommand("PED001", "USR001", SupervisorAutorizado: true));

        Assert.True(result.EsExitoso);
        mock.Verify(r => r.ConfirmarEntregaAsync("PED001", "USR001", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ConfirmarEntrega_PedidoPagado_RetornaOk()
    {
        // Pedido pagado → no requiere supervisor
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED002", It.IsAny<CancellationToken>())).ReturnsAsync(false);
        mock.Setup(r => r.ObtenerEstadoPagoAsync("PED002", It.IsAny<CancellationToken>())).ReturnsAsync("PAGADO");
        mock.Setup(r => r.ConfirmarEntregaAsync("PED002", "USR001", It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var handler = new ConfirmarEntregaCentralHandler(mock.Object);
        var result = await handler.HandleAsync(new ConfirmarEntregaCentralCommand("PED002", "USR001"));

        Assert.True(result.EsExitoso);
        mock.Verify(r => r.ConfirmarEntregaAsync("PED002", "USR001", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ConfirmarEntrega_PedidoYaEntregado_RetornaFallo()
    {
        // BR-DEL-012: pedido ya entregado → no se puede confirmar de nuevo
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED003", It.IsAny<CancellationToken>())).ReturnsAsync(true);

        var handler = new ConfirmarEntregaCentralHandler(mock.Object);
        var result = await handler.HandleAsync(new ConfirmarEntregaCentralCommand("PED003", "USR001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_YA_ENTREGADO", result.CodigoError);
        mock.Verify(r => r.ConfirmarEntregaAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // ── RevertirEntregaCentralHandler ─────────────────────────────────────────

    [Fact]
    public async Task RevertirEntrega_SinSupervisor_RetornaRequiereSupervisor()
    {
        // Siempre requiere supervisor para revertir
        var mock = new Mock<IPedidoDeliveryRepository>();
        var handler = new RevertirEntregaCentralHandler(mock.Object);

        var result = await handler.HandleAsync(new RevertirEntregaCentralCommand("PED001", "USR001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("REQUIERE_SUPERVISOR_22", result.CodigoError);
        mock.Verify(r => r.RevertirEntregaAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RevertirEntrega_ConSupervisor_PedidoEntregado_RetornaOk()
    {
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync(true);
        mock.Setup(r => r.RevertirEntregaAsync("PED001", "USR001", It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var handler = new RevertirEntregaCentralHandler(mock.Object);
        var result = await handler.HandleAsync(new RevertirEntregaCentralCommand("PED001", "USR001", SupervisorAutorizado: true));

        Assert.True(result.EsExitoso);
        mock.Verify(r => r.RevertirEntregaAsync("PED001", "USR001", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task RevertirEntrega_ConSupervisor_PedidoNoEntregado_RetornaFallo()
    {
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync(false);

        var handler = new RevertirEntregaCentralHandler(mock.Object);
        var result = await handler.HandleAsync(new RevertirEntregaCentralCommand("PED001", "USR001", SupervisorAutorizado: true));

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_NO_ENTREGADO", result.CodigoError);
    }

    // ── ModificarFechaProgramadaDeliveryHandler ────────────────────────────────

    [Fact]
    public async Task ModificarFecha_PedidoPendiente_RetornaOk()
    {
        // BR-DEL-014: pedido no entregado → puede modificar fecha
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync(false);
        mock.Setup(r => r.ModificarFechaProgramadaAsync("PED001", It.IsAny<DateTime>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        var handler = new ModificarFechaProgramadaDeliveryHandler(mock.Object);
        var result = await handler.HandleAsync(new ModificarFechaProgramadaDeliveryCommand("PED001", DateTime.Now.AddHours(2)));

        Assert.True(result.EsExitoso);
        mock.Verify(r => r.ModificarFechaProgramadaAsync("PED001", It.IsAny<DateTime>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ModificarFecha_PedidoEntregado_RetornaFallo()
    {
        // BR-DEL-014: pedido entregado → no se puede modificar
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync(true);

        var handler = new ModificarFechaProgramadaDeliveryHandler(mock.Object);
        var result = await handler.HandleAsync(new ModificarFechaProgramadaDeliveryCommand("PED001", DateTime.Now.AddHours(2)));

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_YA_ENTREGADO", result.CodigoError);
    }

    [Fact]
    public async Task ModificarFecha_FechaVacia_RetornaFallo()
    {
        var mock = new Mock<IPedidoDeliveryRepository>();
        mock.Setup(r => r.EstaEntregadoAsync("PED001", It.IsAny<CancellationToken>())).ReturnsAsync(false);

        var handler = new ModificarFechaProgramadaDeliveryHandler(mock.Object);
        var result = await handler.HandleAsync(new ModificarFechaProgramadaDeliveryCommand("PED001", default));

        Assert.False(result.EsExitoso);
        Assert.Equal("DELIVERY_FECHA_REQUERIDA", result.CodigoError);
    }
}
