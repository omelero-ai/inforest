using Inforest.Application.Pedidos;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Pedidos;

/// <summary>
/// Tests para ObtenerDetallePedidoExtendidoHandler.
/// Legacy: frmDetallePedido.frm — Form_Load.
/// Reglas: BR-PEDIDO-005, BR-PEDIDO-006.
/// </summary>
public class DetallePedidoExtendidoHandlerTests
{
    // ── Helpers ──────────────────────────────────────────────────────────────────

    private static Pedido PedidoAbierto(string codigo = "2608000001") =>
        Pedido.Crear(codigo, "M01", "T001", "USR01", CanalVenta.Local, 2);

    private static ItemPedidoVista ItemBase(string item, decimal subtotal = 50m) =>
        new(
            Item: item,
            CodigoProducto: "PROD01",
            Producto: "Lomo Saltado",
            PrecioOficial: 50m,
            Descuento: 0m,
            PrecioVenta: 50m,
            Cantidad: 1m,
            SubTotal: subtotal,
            EstadoItem: "N",
            Facturado: false,
            Imprime: true,
            TienePropiedad: false,
            TieneObservacion: false,
            Corte: false,
            Observacion: null,
            Documento: null,
            Usuario: "MOZO01",
            FechaEnvio: null,
            Comanda: null);

    private static ComboPedidoVista ComboBase(string item, string itemCombo) =>
        new(
            Item: item,
            ItemCombo: itemCombo,
            CodigoProducto: "COMBO",
            CodigoProductoCombo: "COMPOCOMP01",
            Producto: "Bebida incluida",
            Cantidad: 1m,
            Imprime: true,
            TieneObservacion: false,
            TienePropiedad: false,
            Observacion: null);

    private static (Mock<IPedidoRepository> PedidoRepo, Mock<IPedidoReadRepository> ReadRepo)
        BuildMocks(Pedido? pedido, IReadOnlyList<ItemPedidoVista>? items = null,
                   IReadOnlyList<ComboPedidoVista>? combos = null)
    {
        var pedidoRepo = new Mock<IPedidoRepository>();
        var readRepo   = new Mock<IPedidoReadRepository>();

        pedidoRepo
            .Setup(r => r.ObtenerPorCodigoAsync(It.IsAny<string>(), default))
            .ReturnsAsync(pedido);

        readRepo
            .Setup(r => r.ObtenerDetalleExtendidoAsync(It.IsAny<string>(), default))
            .ReturnsAsync(items ?? []);

        readRepo
            .Setup(r => r.ObtenerCombosAsync(It.IsAny<string>(), default))
            .ReturnsAsync(combos ?? []);

        return (pedidoRepo, readRepo);
    }

    private static ObtenerDetallePedidoExtendidoHandler BuildHandler(
        Mock<IPedidoRepository> pr, Mock<IPedidoReadRepository> rr)
        => new(pr.Object, rr.Object);

    // ── Tests ────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task CodigoVacio_RetornaError_BR_PEDIDO_005()
    {
        var (pr, rr) = BuildMocks(null);
        var handler = BuildHandler(pr, rr);

        var result = await handler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery(""));

        Assert.False(result.EsExitoso);
        Assert.Equal("PEDIDO_CODIGO_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task PedidoInexistente_RetornaError()
    {
        var (pr, rr) = BuildMocks(null);
        var handler = BuildHandler(pr, rr);

        var result = await handler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery("NOEXISTE"));

        Assert.False(result.EsExitoso);
        Assert.Equal("PEDIDO_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task PedidoSinItems_RetornaTotalCero()
    {
        var pedido = PedidoAbierto();
        var (pr, rr) = BuildMocks(pedido, items: [], combos: []);
        var handler = BuildHandler(pr, rr);

        var result = await handler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery(pedido.CodigoPedido));

        Assert.True(result.EsExitoso);
        Assert.Equal(0m, result.Valor!.Total);
        Assert.Empty(result.Valor.Items);
        Assert.Empty(result.Valor.Combos);
    }

    [Fact]
    public async Task PedidoConItems_RetornaTotalSumadoDeSubTotales()
    {
        var pedido = PedidoAbierto();
        var items = new List<ItemPedidoVista>
        {
            ItemBase("001", 50m),
            ItemBase("002", 30m)
        };
        var (pr, rr) = BuildMocks(pedido, items: items);
        var handler = BuildHandler(pr, rr);

        var result = await handler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery(pedido.CodigoPedido));

        Assert.True(result.EsExitoso);
        Assert.Equal(80m, result.Valor!.Total);
        Assert.Equal(2, result.Valor.Items.Count);
    }

    [Fact]
    public async Task PedidoConCombos_RetornaCombosEnResultado()
    {
        var pedido = PedidoAbierto();
        var combos = new List<ComboPedidoVista> { ComboBase("001", "C01"), ComboBase("001", "C02") };
        var (pr, rr) = BuildMocks(pedido, combos: combos);
        var handler = BuildHandler(pr, rr);

        var result = await handler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery(pedido.CodigoPedido));

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Combos.Count);
    }

    [Fact]
    public async Task ResultadoCabecera_MapaCodigoPedidoYNumeroAdultos()
    {
        var pedido = PedidoAbierto("2608000042");
        var (pr, rr) = BuildMocks(pedido);
        var handler = BuildHandler(pr, rr);

        var result = await handler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery("2608000042"));

        Assert.True(result.EsExitoso);
        Assert.Equal("2608000042", result.Valor!.CodigoPedido);
        Assert.Equal(2, result.Valor.NumeroAdultos);
    }

    [Fact]
    public async Task ItemsYCombos_AmbasSuitesCargadas_BR_PEDIDO_006()
    {
        var pedido = PedidoAbierto();
        var items  = new List<ItemPedidoVista> { ItemBase("001") };
        var combos = new List<ComboPedidoVista> { ComboBase("001", "C01") };
        var (pr, rr) = BuildMocks(pedido, items: items, combos: combos);
        var handler = BuildHandler(pr, rr);

        var result = await handler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery(pedido.CodigoPedido));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!.Items);
        Assert.Single(result.Valor.Combos);
    }

    [Fact]
    public async Task Total_EsSumaDeSubTotalesDeItems()
    {
        var pedido = PedidoAbierto();
        var items = Enumerable.Range(1, 5)
            .Select(i => ItemBase($"00{i}", 20m))
            .ToList();
        var (pr, rr) = BuildMocks(pedido, items: items);
        var handler = BuildHandler(pr, rr);

        var result = await handler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery(pedido.CodigoPedido));

        Assert.True(result.EsExitoso);
        Assert.Equal(100m, result.Valor!.Total);
    }
}
