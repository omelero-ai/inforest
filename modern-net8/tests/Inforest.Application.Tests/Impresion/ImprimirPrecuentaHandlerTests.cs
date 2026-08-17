using Inforest.Application.Impresion;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Impresion;

/// <summary>
/// Tests para ImprimirPrecuentaHandler y ObtenerImpresorasPorCajaHandler.
/// Legacy: frmPrecuentaImpresora.frm. BR-PRECUENTA-001/002/003.
/// </summary>
public class ImprimirPrecuentaHandlerTests
{
    private static Pedido BuildPedido(string codigo = "PED001", string turno = "TUR01")
    {
        var pedido = Pedido.Crear(codigo, "M01", turno, "USR01", CanalVenta.Local);
        pedido.AgregarDetalle(DetallePedido.Crear(codigo, "001", "PROD01", 2m, 10m, 10m));
        return pedido;
    }

    private static ImprimirPrecuentaHandler BuildHandler(
        Mock<IPedidoRepository>? pedidoRepo = null,
        Mock<IImpresoraService>? impresoraService = null)
        => new(
            (pedidoRepo ?? new Mock<IPedidoRepository>()).Object,
            (impresoraService ?? new Mock<IImpresoraService>()).Object);

    // ── ImprimirPrecuentaHandler ───────────────────────────────────────────────

    [Fact]
    public async Task ImprimirPrecuenta_PedidoYImpresoraValidos_Imprime()
    {
        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildPedido());

        var impresoraService = new Mock<IImpresoraService>();

        var handler = BuildHandler(pedidoRepo, impresoraService);
        var result = await handler.HandleAsync(new ImprimirPrecuentaCommand("PED001", "IMP01"));

        Assert.True(result.EsExitoso);
        impresoraService.Verify(s => s.ImprimirTicketAsync(
            It.Is<string>(c => c.Contains("PRE-CUENTA") && c.Contains("PED001")),
            "IMP01",
            It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ImprimirPrecuenta_PedidoVacio_Falla()
    {
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new ImprimirPrecuentaCommand("", "IMP01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("IMPRESION_PEDIDO_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task ImprimirPrecuenta_ImpresoraVacia_Falla()
    {
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new ImprimirPrecuentaCommand("PED001", ""));

        Assert.False(result.EsExitoso);
        Assert.Equal("IMPRESION_IMPRESORA_REQUERIDA", result.CodigoError);
    }

    [Fact]
    public async Task ImprimirPrecuenta_PedidoNoExiste_Falla()
    {
        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED_INEXISTENTE", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Pedido?)null);

        var handler = BuildHandler(pedidoRepo);
        var result = await handler.HandleAsync(new ImprimirPrecuentaCommand("PED_INEXISTENTE", "IMP01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("IMPRESION_PEDIDO_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task ImprimirPrecuenta_PedidoSinDetalles_Falla()
    {
        var pedido = Pedido.Crear("PED002", "M01", "TUR01", "USR01", CanalVenta.Local);
        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED002", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var handler = BuildHandler(pedidoRepo);
        var result = await handler.HandleAsync(new ImprimirPrecuentaCommand("PED002", "IMP01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("IMPRESION_PEDIDO_SIN_DETALLES", result.CodigoError);
    }

    [Fact]
    public async Task ImprimirPrecuenta_PredeterminadaSinCodigo_Falla()
    {
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new ImprimirPrecuentaCommand(
            "PED001",
            string.Empty,
            UsarImpresoraPredeterminada: true,
            CodigoImpresoraPredeterminada: null));

        Assert.False(result.EsExitoso);
        Assert.Equal("IMPRESION_PRECUENTA_PREDETERMINADA_REQUERIDA", result.CodigoError);
    }

    [Fact]
    public async Task ImprimirPrecuenta_PredeterminadaValida_ImprimeConCodigoPredeterminado()
    {
        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildPedido());

        var impresoraService = new Mock<IImpresoraService>();
        var handler = BuildHandler(pedidoRepo, impresoraService);

        var result = await handler.HandleAsync(new ImprimirPrecuentaCommand(
            "PED001",
            string.Empty,
            UsarImpresoraPredeterminada: true,
            CodigoImpresoraPredeterminada: "PRE001"));

        Assert.True(result.EsExitoso);
        impresoraService.Verify(s => s.ImprimirTicketAsync(
            It.IsAny<string>(),
            "PRE001",
            It.IsAny<CancellationToken>()), Times.Once);
    }

    // ── ObtenerImpresorasPorCajaHandler ───────────────────────────────────────

    [Fact]
    public async Task ObtenerImpresorasPorCaja_CajaValida_RetornaImpresoras()
    {
        var repo = new Mock<IImpresoraRepository>();
        repo.Setup(r => r.ObtenerImpresorasPorCajaAsync("001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { new ImpresoraDisponible("IMP01", "Impresora Caja") });

        var handler = new ObtenerImpresorasPorCajaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerImpresorasPorCajaQuery("001"));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!);
    }

    [Fact]
    public async Task ObtenerImpresorasPorCaja_CajaVacia_Falla()
    {
        var repo = new Mock<IImpresoraRepository>();
        var handler = new ObtenerImpresorasPorCajaHandler(repo.Object);

        var result = await handler.HandleAsync(new ObtenerImpresorasPorCajaQuery(""));

        Assert.False(result.EsExitoso);
        Assert.Equal("IMPRESION_CAJA_REQUERIDA", result.CodigoError);
    }
}
