using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Delivery;

/// <summary>
/// Tests de las entidades de Delivery.
/// Legacy: TDELIVERY, MPEDIDO, TLOCAL, TTIPOPEDIDODETALLE
/// </summary>
public class DeliveryDomainTests
{
    // ── ClienteDelivery ────────────────────────────────────────────────────────

    [Fact]
    public void ClienteDelivery_Crear_ConDatosValidos_CreaCorrectamente()
    {
        var cliente = ClienteDelivery.Crear("DEL001", "01", "García", "Juan", "987654321", "Av. Lima 123", "Z01", "D01");

        Assert.Equal("DEL001", cliente.CodigoDelivery);
        Assert.True(cliente.Activo);
    }

    [Fact]
    public void ClienteDelivery_Crear_SinCodigo_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            ClienteDelivery.Crear("", null, "García", "Juan", "987654321", null, null, null));

        Assert.Equal("DELIVERY_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void ClienteDelivery_Crear_SinContacto_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            ClienteDelivery.Crear("DEL001", null, null, null, null, null, null, null));

        Assert.Equal("DELIVERY_CONTACTO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void ClienteDelivery_AsignarIdentidad_SetValues()
    {
        var cliente = ClienteDelivery.Crear("DEL001", null, "García", null, null, null, null, null);
        cliente.AsignarIdentidad("DNI", "12345678");

        Assert.Equal("DNI", cliente.TipoIdentidad);
        Assert.Equal("12345678", cliente.NumeroIdentidad);
    }

    // ── PedidoDelivery ─────────────────────────────────────────────────────────

    [Fact]
    public void PedidoDelivery_Crear_ConDatosValidos_EstadoPendiente()
    {
        var pedido = PedidoDelivery.Crear("PED001", "TUR001", "USR001", "CAJA001", "DEL001", "Z01", "D01");

        Assert.Equal("PED001", pedido.CodigoPedido);
        Assert.Equal(EstadoDelivery.Pendiente, pedido.EstadoDelivery);
    }

    [Fact]
    public void PedidoDelivery_Crear_SinCliente_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            PedidoDelivery.Crear("PED001", "TUR001", "USR001", "CAJA001", null, null, null));

        Assert.Equal("PEDIDO_DELIVERY_CLIENTE_REQUERIDO", ex.Code);
    }

    [Fact]
    public void PedidoDelivery_AsignarMotorizado_CambiaEstadoEnCamino()
    {
        var pedido = PedidoDelivery.Crear("PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);
        pedido.AsignarMotorizado("MOT001");

        Assert.Equal("MOT001", pedido.CodigoMotorizado);
        Assert.Equal(EstadoDelivery.EnCamino, pedido.EstadoDelivery);
        Assert.NotNull(pedido.FechaDespacho);
    }

    [Fact]
    public void PedidoDelivery_RegistrarEntrega_CambiaEstadoEntregado()
    {
        var pedido = PedidoDelivery.Crear("PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);
        pedido.AsignarMotorizado("MOT001");
        pedido.RegistrarEntrega();

        Assert.Equal(EstadoDelivery.Entregado, pedido.EstadoDelivery);
        Assert.NotNull(pedido.FechaEntrega);
    }

    [Fact]
    public void PedidoDelivery_RegistrarEntrega_SinEstarEnCamino_LanzaExcepcion()
    {
        var pedido = PedidoDelivery.Crear("PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);

        var ex = Assert.Throws<DomainException>(() => pedido.RegistrarEntrega());

        Assert.Equal("PEDIDO_DELIVERY_NO_EN_CAMINO", ex.Code);
    }

    // ── TipoPedidoDelivery ─────────────────────────────────────────────────────

    [Fact]
    public void TipoPedidoDelivery_ValidarPedido_ObligaMotorizado_SinMotorizado_LanzaExcepcion()
    {
        var tipo = TipoPedidoDelivery.Crear("02", true, true, true, false, false);

        var ex = Assert.Throws<DomainException>(() =>
            tipo.ValidarPedidoDelivery("DEL001", null, null));

        Assert.Equal("DELIVERY_MOTORIZADO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void TipoPedidoDelivery_ValidarPedido_ObligaClienteFrecuente_SinCliente_LanzaExcepcion()
    {
        var tipo = TipoPedidoDelivery.Crear("02", true, false, true, false, true);

        var ex = Assert.Throws<DomainException>(() =>
            tipo.ValidarPedidoDelivery(null, "MOT001", null));

        Assert.Equal("DELIVERY_CLIENTE_FRECUENTE_REQUERIDO", ex.Code);
    }

    [Fact]
    public void TipoPedidoDelivery_ValidarPedido_TodosPresentes_NoLanzaExcepcion()
    {
        var tipo = TipoPedidoDelivery.Crear("02", true, true, true, false, true, true);

        var ex = Record.Exception(() =>
            tipo.ValidarPedidoDelivery("DEL001", "MOT001", DateTime.Now.AddHours(1)));

        Assert.Null(ex);
    }

    // ── LocalRestaurante ───────────────────────────────────────────────────────

    [Fact]
    public void LocalRestaurante_Crear_ConCodigo_CreaCorrectamente()
    {
        var local = LocalRestaurante.Crear("001", "Local Principal", "192.168.1.1", "INFOREST", true, true);

        Assert.Equal("001", local.CodigoLocal);
        Assert.True(local.Activo);
        Assert.True(local.ParticipanEnReplica);
    }

    [Fact]
    public void LocalRestaurante_Crear_SinCodigo_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            LocalRestaurante.Crear("", null, null, null, true, false));

        Assert.Equal("LOCAL_CODIGO_REQUERIDO", ex.Code);
    }
}
