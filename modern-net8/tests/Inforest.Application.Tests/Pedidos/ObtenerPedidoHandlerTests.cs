using Inforest.Application.Pedidos;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Pedidos;

/// <summary>
/// Tests de los handlers de consulta de pedidos.
/// Legacy: vPedido view, MPEDIDO, DPEDIDO. BR-SQL-001/BR-SQL-002.
/// </summary>
public class ObtenerPedidoHandlerTests
{
    // ---- ObtenerPedidoPorCodigoHandler ----

    [Fact]
    public async Task ObtenerPorCodigo_PedidoExiste_RetornaOk()
    {
        var pedido = Pedido.Crear("PED001", "M01", "TURNO01", "USR01", CanalVenta.Local);
        var repo = new Mock<IPedidoRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var handler = new ObtenerPedidoPorCodigoHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerPedidoPorCodigoQuery("PED001"));

        Assert.True(result.EsExitoso);
        Assert.NotNull(result.Valor);
        Assert.Equal("PED001", result.Valor!.CodigoPedido);
    }

    [Fact]
    public async Task ObtenerPorCodigo_PedidoNoExiste_RetornaOkConNull()
    {
        var repo = new Mock<IPedidoRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("NOEXISTE", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Pedido?)null);

        var handler = new ObtenerPedidoPorCodigoHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerPedidoPorCodigoQuery("NOEXISTE"));

        Assert.True(result.EsExitoso);
        Assert.Null(result.Valor);
    }

    // ---- ObtenerPedidosPorTurnoHandler ----

    [Fact]
    public async Task ObtenerPorTurno_ConPedidos_RetornaLista()
    {
        var pedidos = new List<Pedido>
        {
            Pedido.Crear("PED001", "M01", "TURNO01", "USR01", CanalVenta.Local),
            Pedido.Crear("PED002", "M02", "TURNO01", "USR01", CanalVenta.Delivery)
        };
        var readRepo = new Mock<IPedidoReadRepository>();
        readRepo.Setup(x => x.ObtenerPorTurnoAsync("TURNO01", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedidos.AsReadOnly());

        var handler = new ObtenerPedidosPorTurnoHandler(readRepo.Object);
        var result = await handler.HandleAsync(new ObtenerPedidosPorTurnoQuery("TURNO01"));

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count);
    }

    [Fact]
    public async Task ObtenerPorTurno_SinPedidos_RetornaListaVacia()
    {
        var readRepo = new Mock<IPedidoReadRepository>();
        readRepo.Setup(x => x.ObtenerPorTurnoAsync("TURNO99", It.IsAny<CancellationToken>()))
            .ReturnsAsync(Array.Empty<Pedido>());

        var handler = new ObtenerPedidosPorTurnoHandler(readRepo.Object);
        var result = await handler.HandleAsync(new ObtenerPedidosPorTurnoQuery("TURNO99"));

        Assert.True(result.EsExitoso);
        Assert.Empty(result.Valor!);
    }

    // ---- ObtenerPedidosPorMesaHandler ----

    [Fact]
    public async Task ObtenerPorMesa_ConPedidos_RetornaLista()
    {
        var pedidos = new List<Pedido>
        {
            Pedido.Crear("PED001", "M01", "TURNO01", "USR01", CanalVenta.Local)
        };
        var readRepo = new Mock<IPedidoReadRepository>();
        readRepo.Setup(x => x.ObtenerPorMesaAsync("M01", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedidos.AsReadOnly());

        var handler = new ObtenerPedidosPorMesaHandler(readRepo.Object);
        var result = await handler.HandleAsync(new ObtenerPedidosPorMesaQuery("M01"));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!);
        Assert.Equal("M01", result.Valor![0].CodigoMesa);
    }

    [Fact]
    public async Task ObtenerPorMesa_SinPedidos_RetornaListaVacia()
    {
        var readRepo = new Mock<IPedidoReadRepository>();
        readRepo.Setup(x => x.ObtenerPorMesaAsync("M99", It.IsAny<CancellationToken>()))
            .ReturnsAsync(Array.Empty<Pedido>());

        var handler = new ObtenerPedidosPorMesaHandler(readRepo.Object);
        var result = await handler.HandleAsync(new ObtenerPedidosPorMesaQuery("M99"));

        Assert.True(result.EsExitoso);
        Assert.Empty(result.Valor!);
    }
}
