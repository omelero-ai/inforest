using Inforest.Domain.Entities.Productos;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Services;

namespace Inforest.Domain.Tests.Entities;

/// <summary>
/// Tests de la entidad Pedido y reglas de negocio relacionadas.
/// Legacy: MPEDIDO, DPEDIDO, modPuntoVenta.bas, frmVenta.frm
/// </summary>
public class PedidoTests
{
    [Fact]
    public void Crear_ConDatosValidos_CreaConEstadoAbierto()
    {
        var pedido = Pedido.Crear("PED001", "MESA01", "TURNO001", "USR001", CanalVenta.Local);

        Assert.Equal("PED001", pedido.CodigoPedido);
        Assert.Equal(EstadoPedido.Abierto, pedido.Estado);
        Assert.Equal(CanalVenta.Local, pedido.CanalVenta);
    }

    [Fact]
    public void Crear_SinCodigoPedido_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            Pedido.Crear("", "MESA01", "TURNO001", "USR001", CanalVenta.Local));

        Assert.Equal("PEDIDO_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Crear_SinTurno_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            Pedido.Crear("PED001", "MESA01", "", "USR001", CanalVenta.Local));

        Assert.Equal("PEDIDO_TURNO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Anular_PedidoAbierto_CambiaEstadoAnulado()
    {
        var pedido = Pedido.Crear("PED001", null, "TURNO001", "USR001", CanalVenta.Delivery);
        pedido.Anular("motivo", "USR001");

        Assert.Equal(EstadoPedido.Anulado, pedido.Estado);
    }
}

/// <summary>
/// Tests del servicio de precios.
/// Legacy: BR-001 — modDeclaracion.bas + TPRODUCTO precios por canal
/// </summary>
public class PrecioServiceTests
{
    private readonly PrecioService _service = new();

    [Fact]
    public void CalcularPrecioNeto_SinImpuesto_RetornaMismoPrecio()
    {
        var neto = _service.CalcularPrecioNeto(100m, 0, 0, 0);
        Assert.Equal(100m, neto);
    }

    [Fact]
    public void CalcularPrecioNeto_Con18PorcentoIGV_RetornaNetoCorreecto()
    {
        // Perú: IGV 18% — precio venta 118 → neto = 118 / 1.18 = 100
        var neto = _service.CalcularPrecioNeto(118m, 18m, 0, 0);
        Assert.Equal(100m, Math.Round(neto, 2));
    }

    [Fact]
    public void CalcularTotalLinea_CantidadPorPrecio_RetornaTotalCorrecto()
    {
        var total = _service.CalcularTotalLinea(3m, 25.50m);
        Assert.Equal(76.50m, total);
    }
}
