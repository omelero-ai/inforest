using Inforest.Domain.Entities.Ventas;

namespace Inforest.Domain.Tests.Pedidos;

/// <summary>
/// Tests de dominio para la División de Pedidos.
/// Reglas: BR-DIV-001 a BR-DIV-005.
/// Legacy: frmDivision.frm.
/// </summary>
public class DivisionPedidoTests
{
    // ─── Helpers ────────────────────────────────────────────────────────────────

    private static ItemDivision CrearItem(string item, double cantidad, double venta = 0, double imp1 = 0)
        => new(
            item: item,
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
            impuesto1: imp1 > 0 ? imp1 : 1.8 * cantidad,
            impuesto2: 0,
            impuesto3: 0,
            venta: venta > 0 ? venta : 11.8 * cantidad,
            observacion: null,
            cortesia: null,
            imprime: true,
            estadoItem: null,
            area: null);

    private static SesionDivision CrearSesion(IEnumerable<ItemDivision> items)
        => SesionDivision.Crear("2608000001", "01", items);

    // ─── Tests AgregarPedidoDestino ──────────────────────────────────────────────

    [Fact]
    public void AgregarPedidoDestino_AgregaPedidoConCodigoTemporal()
    {
        var sesion = CrearSesion([CrearItem("001", 1)]);
        var pedido = sesion.AgregarPedidoDestino("12345", "Mesa 1 - 001");
        Assert.Single(sesion.PedidosDestino);
        Assert.Equal("001", pedido.CodigoTemporal);
    }

    [Fact]
    public void AgregarMultiplesPedidosDestino_CodigosSecuenciales()
    {
        var sesion = CrearSesion([CrearItem("001", 1)]);
        sesion.AgregarPedidoDestino("0", "obs1");
        sesion.AgregarPedidoDestino("0", "obs2");
        sesion.AgregarPedidoDestino("0", "obs3");
        Assert.Equal(3, sesion.PedidosDestino.Count);
        Assert.Equal("001", sesion.PedidosDestino[0].CodigoTemporal);
        Assert.Equal("002", sesion.PedidosDestino[1].CodigoTemporal);
        Assert.Equal("003", sesion.PedidosDestino[2].CodigoTemporal);
    }

    // ─── Tests Disgregar ─────────────────────────────────────────────────────────

    [Fact]
    public void Disgregar_CantidadTotal_CreaItemsIndividuales_BR_DIV_003()
    {
        // Item con cantidad=3, venta=35.40 → 3 items con cantidad=1, venta=11.80 c/u
        var sesion = CrearSesion([CrearItem("001", 3, venta: 35.4, imp1: 5.4)]);
        bool ok = sesion.Disgregar("001");
        Assert.True(ok);
        Assert.Equal(3, sesion.Items.Count);
        Assert.All(sesion.Items, i => Assert.Equal(1, i.Cantidad));
        Assert.All(sesion.Items, i => Assert.Equal(11.8, i.Venta, 4));
    }

    [Fact]
    public void Disgregar_ConCantidadEspecifica_CreaDosItems_BR_DIV_003()
    {
        var sesion = CrearSesion([CrearItem("001", 5, venta: 59.0, imp1: 9.0)]);
        bool ok = sesion.Disgregar("001", cantidadPrimera: 2);
        Assert.True(ok);
        Assert.Equal(2, sesion.Items.Count);
        var first = sesion.Items[0];
        var second = sesion.Items[1];
        Assert.Equal(2, first.Cantidad);
        Assert.Equal(3, second.Cantidad);
        // venta proporcional
        Assert.Equal(23.6, first.Venta, 2);
        Assert.Equal(35.4, second.Venta, 2);
    }

    [Fact]
    public void Disgregar_CantidadUno_RetornaFalse()
    {
        var sesion = CrearSesion([CrearItem("001", 1)]);
        bool ok = sesion.Disgregar("001");
        Assert.False(ok);
    }

    [Fact]
    public void Disgregar_CantidadInvalida_RetornaFalse()
    {
        var sesion = CrearSesion([CrearItem("001", 3)]);
        bool ok = sesion.Disgregar("001", cantidadPrimera: 0);
        Assert.False(ok);
        // cantidad igual tampoco
        bool ok2 = sesion.Disgregar("001", cantidadPrimera: 3);
        Assert.False(ok2);
    }

    // ─── Tests MoverItem ─────────────────────────────────────────────────────────

    [Fact]
    public void MoverItemADestino_AsignaCodigoPedidoDestino_BR_DIV_005()
    {
        var sesion = CrearSesion([CrearItem("001", 1), CrearItem("002", 2)]);
        sesion.AgregarPedidoDestino("0", "obs");
        var codigoDestino = sesion.PedidosDestino[0].CodigoTemporal;

        bool ok = sesion.MoverItemADestino("001", codigoDestino);
        Assert.True(ok);
        Assert.Single(sesion.ItemsOrigen);
        Assert.Single(sesion.ItemsDestino(codigoDestino));
    }

    [Fact]
    public void RevertirItemAOrigen_RegresaItemAlOrigen_BR_DIV_005()
    {
        var sesion = CrearSesion([CrearItem("001", 1)]);
        sesion.AgregarPedidoDestino("0", "obs");
        var cod = sesion.PedidosDestino[0].CodigoTemporal;
        sesion.MoverItemADestino("001", cod);

        bool ok = sesion.RevertirItemAOrigen("001", cod);
        Assert.True(ok);
        Assert.Single(sesion.ItemsOrigen);
        Assert.Empty(sesion.ItemsDestino(cod));
    }

    // ─── Tests Compartir ─────────────────────────────────────────────────────────

    [Fact]
    public void Compartir_DivideEquitativamente_BR_DIV_004()
    {
        // 1 item venta=35.4 → 2 pedidos destino + origen = 3 → cada uno venta=11.8
        var sesion = CrearSesion([CrearItem("001", 3, venta: 35.4, imp1: 5.4)]);
        sesion.AgregarPedidoDestino("0", "obs1");
        sesion.AgregarPedidoDestino("0", "obs2");

        bool ok = sesion.Compartir();
        Assert.True(ok);
        // Origen ahora tiene 1 item con cantidad=1 y los destinos también
        var origen = sesion.ItemsOrigen.First();
        Assert.Equal(1, origen.Cantidad);
        Assert.Equal(11.8, origen.Venta, 2);
        // Cada destino tiene 1 item
        foreach (var p in sesion.PedidosDestino)
            Assert.Single(sesion.ItemsDestino(p.CodigoTemporal));
    }

    [Fact]
    public void Compartir_SinPedidosDestino_RetornaFalse_BR_DIV_004()
    {
        var sesion = CrearSesion([CrearItem("001", 2)]);
        bool ok = sesion.Compartir();
        Assert.False(ok);
    }

    // ─── Tests MontoOrigen ───────────────────────────────────────────────────────

    [Fact]
    public void MontoOrigen_SumaCorrectamente()
    {
        var sesion = CrearSesion([CrearItem("001", 1, venta: 10), CrearItem("002", 2, venta: 20)]);
        Assert.Equal(30m, sesion.MontoOrigen);
    }

    [Fact]
    public void MontoOrigen_ExcluyeItemsMovidos()
    {
        var sesion = CrearSesion([CrearItem("001", 1, venta: 10), CrearItem("002", 2, venta: 20)]);
        sesion.AgregarPedidoDestino("0", "obs");
        sesion.MoverItemADestino("002", sesion.PedidosDestino[0].CodigoTemporal);
        Assert.Equal(10m, sesion.MontoOrigen);
    }
}
