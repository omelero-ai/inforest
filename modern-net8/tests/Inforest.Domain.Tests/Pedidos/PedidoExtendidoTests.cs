using System.Reflection;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Pedidos;

public class PedidoExtendidoTests
{
    [Fact]
    public void CrearPedidoConItems_AgregaDetalleCorrectamente()
    {
        var pedido = Pedido.Crear("PED001", "M01", "TURNO01", "USR01", CanalVenta.Local, 2);
        var detalle = DetallePedido.Crear("PED001", "001", "PRD01", 2m, 10m, 10m);

        pedido.AgregarDetalle(detalle);

        Assert.Single(pedido.Detalles);
        Assert.Equal(20m, pedido.Detalles[0].Venta);
    }

    [Fact]
    public void AgregarItemAPedidoCerrado_LanzaExcepcion()
    {
        var pedido = Pedido.Crear("PED001", "M01", "TURNO01", "USR01", CanalVenta.Local, 2);
        var detalle = DetallePedido.Crear("PED001", "001", "PRD01", 1m, 10m, 10m);
        typeof(Pedido).GetProperty(nameof(Pedido.Estado), BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)!
            .SetValue(pedido, EstadoPedido.Cerrado);

        var ex = Assert.Throws<DomainException>(() => pedido.AgregarDetalle(detalle));
        Assert.Equal("PEDIDO_ESTADO_INVALIDO", ex.Code);
    }
}
