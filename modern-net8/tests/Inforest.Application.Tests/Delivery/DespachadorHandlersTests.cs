using Inforest.Application.Configuracion;
using Inforest.Application.Despacho;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Entities.Motorizado;
using Inforest.Domain.Repositories;
using Moq;
using DomainMotorizado = Inforest.Domain.Entities.Motorizado.Motorizado;

namespace Inforest.Application.Tests.Delivery;

public sealed class DespachadorHandlersTests
{
    [Fact]
    public async Task AsignarMotorizado_SinEmpacador_RetornaError()
    {
        var pedidos = new Mock<IPedidoDeliveryRepository>();
        pedidos.Setup(r => r.ObtenerResumenDespachadorPorPedidoAsync("P001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new PedidoDespachadorResumen("P001", DateTime.Today, "USR", "001", "CLIENTE", "999", "DIR", "Z", "", "0000", null, 10));

        var motorizados = new Mock<IMotorizadoRepository>();
        var parametro = new Mock<IParametroRepository>();

        var handler = new AsignarMotorizadoDespachoHandler(pedidos.Object, motorizados.Object, parametro.Object);
        var result = await handler.HandleAsync(new AsignarMotorizadoDespachoCommand("P001", "M001", "USR"));

        Assert.False(result.EsExitoso);
        Assert.Equal("DESPACHADOR_REQUIERE_EMPACADOR", result.CodigoError);
    }

    [Fact]
    public async Task AsignarMotorizado_PedidoYaAsignado_RetornaError()
    {
        var pedidos = new Mock<IPedidoDeliveryRepository>();
        pedidos.Setup(r => r.ObtenerResumenDespachadorPorPedidoAsync("P001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new PedidoDespachadorResumen("P001", DateTime.Today, "USR", "001", "CLIENTE", "999", "DIR", "Z", "EMP", "M001", DateTime.Now, 10));

        var motorizados = new Mock<IMotorizadoRepository>();
        var parametro = new Mock<IParametroRepository>();

        var handler = new AsignarMotorizadoDespachoHandler(pedidos.Object, motorizados.Object, parametro.Object);
        var result = await handler.HandleAsync(new AsignarMotorizadoDespachoCommand("P001", "M001", "USR"));

        Assert.False(result.EsExitoso);
        Assert.Equal("DESPACHADOR_PEDIDO_YA_ASIGNADO", result.CodigoError);
    }

    [Fact]
    public async Task AsignarMotorizado_Valido_PersistesAsignacion()
    {
        var pedidos = new Mock<IPedidoDeliveryRepository>();
        pedidos.Setup(r => r.ObtenerResumenDespachadorPorPedidoAsync("P001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new PedidoDespachadorResumen("P001", DateTime.Today, "USR", "001", "CLIENTE", "999", "DIR", "Z", "EMP001", "0000", null, 10));
        pedidos.Setup(r => r.ContarAsignacionesPrincipalesMotorizadoAsync("M001", It.IsAny<DateTime>(), It.IsAny<CancellationToken>())).ReturnsAsync(0);
        pedidos.Setup(r => r.ContarReasignacionesMotorizadoAsync("M001", It.IsAny<DateTime>(), It.IsAny<CancellationToken>())).ReturnsAsync(0);
        pedidos.Setup(r => r.AsignarMotorizadoDespachoAsync("P001", "M001", It.IsAny<decimal>(), false, It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var motorizados = new Mock<IMotorizadoRepository>();
        motorizados.Setup(r => r.ObtenerPorCodigoAsync("M001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(DomainMotorizado.Crear("M001", "DNI", 5, 7, 9));

        var parametro = new Mock<IParametroRepository>();
        parametro.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>())).ReturnsAsync((Inforest.Domain.Entities.Configuracion.ConfiguracionSistema?)null);

        var handler = new AsignarMotorizadoDespachoHandler(pedidos.Object, motorizados.Object, parametro.Object);
        var result = await handler.HandleAsync(new AsignarMotorizadoDespachoCommand("P001", "M001", "USR"));

        Assert.True(result.EsExitoso);
        pedidos.Verify(r => r.AsignarMotorizadoDespachoAsync("P001", "M001", 5, false, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task DesasignarMotorizado_NoAsignado_RetornaError()
    {
        var pedidos = new Mock<IPedidoDeliveryRepository>();
        pedidos.Setup(r => r.ObtenerResumenDespachadorPorPedidoAsync("P001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new PedidoDespachadorResumen("P001", DateTime.Today, "USR", "001", "CLIENTE", "999", "DIR", "Z", "EMP", "0000", null, 10));

        var handler = new DesasignarMotorizadoDespachoHandler(pedidos.Object);
        var result = await handler.HandleAsync(new DesasignarMotorizadoDespachoCommand("P001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("DESPACHADOR_PEDIDO_NO_ASIGNADO", result.CodigoError);
    }

    [Fact]
    public async Task DesasignarEmpacador_ConMotorizado_RetornaError()
    {
        var pedidos = new Mock<IPedidoDeliveryRepository>();
        pedidos.Setup(r => r.ObtenerResumenDespachadorPorPedidoAsync("P001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new PedidoDespachadorResumen("P001", DateTime.Today, "USR", "001", "CLIENTE", "999", "DIR", "Z", "EMP001", "M001", DateTime.Now, 10));

        var handler = new DesasignarEmpacadorDespachoHandler(pedidos.Object);
        var result = await handler.HandleAsync(new DesasignarEmpacadorDespachoCommand("P001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("DESPACHADOR_PEDIDO_EN_CAMINO", result.CodigoError);
    }
}
