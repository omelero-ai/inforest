using Inforest.Application.Interfaces;
using Inforest.Application.Pedidos;
using Inforest.Domain.Entities.Ventas;
using Moq;

namespace Inforest.Application.Tests.Pedidos;

/// <summary>
/// Tests de handlers para la División de Pedidos.
/// Legacy: frmDivision.frm — cmdOpcion Aceptar + iniciar sesión.
/// </summary>
public class DivisionPedidoHandlerTests
{
    private static ItemDivisionRecord RecordBase(string item, double cantidad = 1, double venta = 11.8)
        => new(
            CodigoPedido: "2608000001",
            Item: item,
            TipoPedido: "01",
            CodigoProducto: "PROD01",
            NombreProducto: "Producto Test",
            CodigoGrupo: "G1",
            CodigoSubGrupo: "S1",
            Moneda: "SOL",
            PrecioNeto: 10,
            PrecioImpuesto1: 1.8,
            PrecioImpuesto2: 0,
            PrecioImpuesto3: 0,
            PrecioVenta: 11.8,
            Recargo: 0,
            Descuento: 0,
            PrecioOficial: 11.8,
            Cantidad: cantidad,
            Impuesto1: 1.8 * cantidad,
            Impuesto2: 0,
            Impuesto3: 0,
            Venta: venta * cantidad,
            Observacion: null,
            Cortesia: null,
            Imprime: true,
            EstadoItem: null,
            Area: null);

    // ─── IniciarDivisionHandler ──────────────────────────────────────────────────

    [Fact]
    public async Task IniciarDivision_EstadoEmitido_RetornaSesion_BR_DIV_001()
    {
        var repo = new Mock<IDivisionPedidoRepository>();
        repo.Setup(r => r.ObtenerEstadoPedidoAsync("PED001", default)).ReturnsAsync("01");
        repo.Setup(r => r.ObtenerTipoPedidoAsync("PED001", default)).ReturnsAsync("01");
        repo.Setup(r => r.CargarItemsPedidoAsync("PED001", default))
            .ReturnsAsync([RecordBase("001"), RecordBase("002")]);

        var handler = new IniciarDivisionHandler(repo.Object);
        var result = await handler.HandleAsync(new IniciarDivisionCommand("PED001", "CAJ1", "12345", "Mesa 1"));

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Items.Count);
    }

    [Fact]
    public async Task IniciarDivision_EstadoDiferenteEmitido_RetornaError_BR_DIV_001()
    {
        var repo = new Mock<IDivisionPedidoRepository>();
        repo.Setup(r => r.ObtenerEstadoPedidoAsync("PED001", default)).ReturnsAsync("03"); // cerrado

        var handler = new IniciarDivisionHandler(repo.Object);
        var result = await handler.HandleAsync(new IniciarDivisionCommand("PED001", "CAJ1", "0", "obs"));

        Assert.False(result.EsExitoso);
        Assert.Contains("03", result.MensajeError);
    }

    [Fact]
    public async Task IniciarDivision_EstadoNulo_RetornaError()
    {
        var repo = new Mock<IDivisionPedidoRepository>();
        repo.Setup(r => r.ObtenerEstadoPedidoAsync("PED001", default)).ReturnsAsync((string?)null);

        var handler = new IniciarDivisionHandler(repo.Object);
        var result = await handler.HandleAsync(new IniciarDivisionCommand("PED001", "CAJ1", "0", "obs"));

        Assert.False(result.EsExitoso);
    }

    // ─── ConfirmarDivisionHandler ────────────────────────────────────────────────

    [Fact]
    public async Task ConfirmarDivision_SinPedidosDestino_LlamaConfirmarConSoloOrigen()
    {
        var repo = new Mock<IDivisionPedidoRepository>();
        repo.Setup(r => r.ObtenerEstadoPedidoAsync("2608000001", default)).ReturnsAsync("01");

        var handler = new ConfirmarDivisionHandler(repo.Object);
        var sesion = CrearSesionConItemsBase();
        var cmd = new ConfirmarDivisionCommand(sesion, "20260817");

        var result = await handler.HandleAsync(cmd);

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.ConfirmarDivisionAsync(
            It.Is<ConfirmarDivisionParams>(p => p.CodigoPedidoOrigen == "2608000001"),
            default), Times.Once);
    }

    [Fact]
    public async Task ConfirmarDivision_EstadoNoEmitido_RetornaError_BR_DIV_001()
    {
        var repo = new Mock<IDivisionPedidoRepository>();
        repo.Setup(r => r.ObtenerEstadoPedidoAsync(It.IsAny<string>(), default)).ReturnsAsync("AN");

        var handler = new ConfirmarDivisionHandler(repo.Object);
        var sesion = CrearSesionConItemsBase();
        var result = await handler.HandleAsync(new ConfirmarDivisionCommand(sesion, "20260817"));

        Assert.False(result.EsExitoso);
        repo.Verify(r => r.ConfirmarDivisionAsync(It.IsAny<ConfirmarDivisionParams>(), default), Times.Never);
    }

    [Fact]
    public async Task ConfirmarDivision_MontoMaximoSuperado_RetornaError_BR_DIV_002()
    {
        var repo = new Mock<IDivisionPedidoRepository>();
        repo.Setup(r => r.ObtenerEstadoPedidoAsync("2608000001", default)).ReturnsAsync("01");
        repo.Setup(r => r.ObtenerMontoMaximoPedidoAsync(It.IsAny<string>(), default)).ReturnsAsync(5m);
        repo.Setup(r => r.ObtenerVentaActualPedidoAsync(It.IsAny<string>(), default)).ReturnsAsync(0m);
        repo.Setup(r => r.GenerarSiguienteCorrelativoPedidoAsync(default)).ReturnsAsync("2600000001");

        var handler = new ConfirmarDivisionHandler(repo.Object);
        var sesion = CrearSesionConItemsBase();
        sesion.AgregarPedidoDestino("0", "obs");
        var cod = sesion.PedidosDestino[0].CodigoTemporal;
        sesion.MoverItemADestino("001", cod); // item con venta=11.8 > máximo 5

        var result = await handler.HandleAsync(new ConfirmarDivisionCommand(sesion, "20260817"));

        Assert.False(result.EsExitoso);
        Assert.Contains("monto máximo", result.MensajeError, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConfirmarDivision_ConPedidoDestinoConItems_LlamaGenerarCorrelativo()
    {
        var repo = new Mock<IDivisionPedidoRepository>();
        repo.Setup(r => r.ObtenerEstadoPedidoAsync("2608000001", default)).ReturnsAsync("01");
        repo.Setup(r => r.ObtenerMontoMaximoPedidoAsync(It.IsAny<string>(), default)).ReturnsAsync(0m);
        repo.Setup(r => r.ObtenerVentaActualPedidoAsync(It.IsAny<string>(), default)).ReturnsAsync(0m);
        repo.Setup(r => r.GenerarSiguienteCorrelativoPedidoAsync(default)).ReturnsAsync("2600000002");

        var handler = new ConfirmarDivisionHandler(repo.Object);
        var sesion = CrearSesionConItemsBase();
        sesion.AgregarPedidoDestino("0", "obs");
        sesion.MoverItemADestino("001", sesion.PedidosDestino[0].CodigoTemporal);

        var result = await handler.HandleAsync(new ConfirmarDivisionCommand(sesion, "20260817"));

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.GenerarSiguienteCorrelativoPedidoAsync(default), Times.Once);
        repo.Verify(r => r.ConfirmarDivisionAsync(It.IsAny<ConfirmarDivisionParams>(), default), Times.Once);
    }

    // ─── DisgregarItemHandler ────────────────────────────────────────────────────

    [Fact]
    public void Disgregar_ItemValido_RetornaExitoso_BR_DIV_003()
    {
        var handler = new DisgregarItemHandler();
        var sesion = CrearSesionConItemsBase(cantidad: 3);
        var result = handler.HandleAsync(new DisgregarItemCommand(sesion, "001", null));
        Assert.True(result.EsExitoso);
    }

    [Fact]
    public void Disgregar_CantidadInvalida_RetornaError()
    {
        var handler = new DisgregarItemHandler();
        var sesion = CrearSesionConItemsBase(cantidad: 3);
        var result = handler.HandleAsync(new DisgregarItemCommand(sesion, "001", 5));
        Assert.False(result.EsExitoso);
    }

    // ─── CompartirPedidoHandler ──────────────────────────────────────────────────

    [Fact]
    public void Compartir_ConPedidosDestino_RetornaExitoso_BR_DIV_004()
    {
        var handler = new CompartirPedidoHandler();
        var sesion = CrearSesionConItemsBase();
        sesion.AgregarPedidoDestino("0", "obs");
        var result = handler.HandleAsync(sesion);
        Assert.True(result.EsExitoso);
    }

    [Fact]
    public void Compartir_SinPedidosDestino_RetornaError()
    {
        var handler = new CompartirPedidoHandler();
        var sesion = CrearSesionConItemsBase();
        var result = handler.HandleAsync(sesion);
        Assert.False(result.EsExitoso);
    }

    // ─── Helpers ─────────────────────────────────────────────────────────────────

    private static SesionDivision CrearSesionConItemsBase(double cantidad = 1)
    {
        var item1 = new ItemDivision(
            item: "001",
            codigoPedido: "2608000001",
            tipoPedido: "01",
            codigoProducto: "PROD01",
            codigoGrupo: "G1",
            codigoSubGrupo: "S1",
            moneda: "SOL",
            precioNeto: 10,
            precioImpuesto1: 1.8,
            precioImpuesto2: 0,
            precioImpuesto3: 0,
            precioVenta: 11.8,
            recargo: 0,
            descuento: 0,
            precioOficial: 11.8,
            cantidad: cantidad,
            impuesto1: 1.8 * cantidad,
            impuesto2: 0,
            impuesto3: 0,
            venta: 11.8 * cantidad,
            observacion: null,
            cortesia: null,
            imprime: true,
            estadoItem: null,
            area: null);
        return SesionDivision.Crear("2608000001", "01", [item1]);
    }
}
