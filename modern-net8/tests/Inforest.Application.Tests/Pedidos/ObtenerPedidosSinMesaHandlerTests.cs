using Inforest.Application.Pedidos;
using Moq;

namespace Inforest.Application.Tests.Pedidos;

/// <summary>
/// Tests del handler de pedidos sin mesa asignada.
/// Legacy: frmMesas.frm — MPEDIDO WHERE tEstadoPedido='01' AND tTipoPedido &lt;&gt; '04' AND LEN(RTRIM(tMesa))=0.
/// Regla BR-MESAS-005.
/// </summary>
public class ObtenerPedidosSinMesaHandlerTests
{
    [Fact]
    public async Task Handle_ConResultados_RetornaListaOk()
    {
        var repo = new Mock<IPedidoReadRepository>();
        var esperado = new List<PedidoSinMesaVista>
        {
            new("PED001", "Delivery zona norte"),
            new("PED002", "")
        };
        repo.Setup(x => x.ObtenerActivosSinMesaAsync("C01", It.IsAny<CancellationToken>()))
            .ReturnsAsync(esperado);

        var handler = new ObtenerPedidosSinMesaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerPedidosSinMesaQuery("C01"));

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count);
        Assert.Equal("PED001", result.Valor![0].CodigoPedido);
        Assert.Equal("Delivery zona norte", result.Valor![0].Observacion);
    }

    [Fact]
    public async Task Handle_SinResultados_RetornaListaVaciaOk()
    {
        var repo = new Mock<IPedidoReadRepository>();
        repo.Setup(x => x.ObtenerActivosSinMesaAsync("C02", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<PedidoSinMesaVista>());

        var handler = new ObtenerPedidosSinMesaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerPedidosSinMesaQuery("C02"));

        Assert.True(result.EsExitoso);
        Assert.Empty(result.Valor!);
    }

    [Fact]
    public async Task Handle_PropagaCajaAlRepositorio()
    {
        var repo = new Mock<IPedidoReadRepository>();
        repo.Setup(x => x.ObtenerActivosSinMesaAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<PedidoSinMesaVista>());

        var handler = new ObtenerPedidosSinMesaHandler(repo.Object);
        await handler.HandleAsync(new ObtenerPedidosSinMesaQuery("CAJA-99"));

        repo.Verify(x => x.ObtenerActivosSinMesaAsync("CAJA-99", It.IsAny<CancellationToken>()), Times.Once);
    }
}
