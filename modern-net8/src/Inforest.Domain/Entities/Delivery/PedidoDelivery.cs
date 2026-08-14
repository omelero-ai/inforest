using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Delivery;

/// <summary>
/// Estado del ciclo de vida de un pedido delivery.
/// <para>
/// Legacy: campo <c>tEstadoPedido</c> en <c>MPEDIDO</c> — valores usados en vistas <c>vDespachador</c> y <c>vDelivery</c>.
/// </para>
/// </summary>
public enum EstadoDelivery
{
    Pendiente = 0,
    EnPreparacion = 1,
    EnCamino = 2,
    Entregado = 3,
    Anulado = 4
}

/// <summary>
/// Cabecera de un pedido con datos extendidos de delivery.
/// <para>
/// Legacy: <c>MPEDIDO</c> con campos adicionales tClienteDelivery, tZona, tDistrito,
/// fHoraEstimadaEntrega y tEstadoDelivery.  Usada por <c>frmDespachador.frm</c>,
/// <c>frmPedidoDelivery.frm</c> y <c>frmNuevoDelivery.frm</c>.
/// </para>
/// Reglas: BR-DEL-001, BR-DEL-009, BR-DEL-010
/// </summary>
public class PedidoDelivery : Entity
{
    public string CodigoPedido { get; private set; } = string.Empty;
    public string CodigoTurno { get; private set; } = string.Empty;
    public string CodigoUsuario { get; private set; } = string.Empty;
    public string CodigoCaja { get; private set; } = string.Empty;
    public string? CodigoClienteDelivery { get; private set; }
    public string? CodigoZona { get; private set; }
    public string? CodigoDistrito { get; private set; }
    public string? CodigoMotorizado { get; private set; }
    public EstadoDelivery EstadoDelivery { get; private set; }
    public DateTime FechaRegistro { get; private set; }
    public DateTime? HoraEstimadaEntrega { get; private set; }
    public DateTime? FechaDespacho { get; private set; }
    public DateTime? FechaEntrega { get; private set; }
    public decimal MontoTotal { get; private set; }
    public string? TipoPedido { get; private set; }
    public string? Observacion { get; private set; }
    public bool EsPrioridad { get; private set; }

    // Dapper constructor
    private PedidoDelivery() { }

    public static PedidoDelivery Crear(
        string codigoPedido,
        string codigoTurno,
        string codigoUsuario,
        string codigoCaja,
        string? codigoClienteDelivery,
        string? codigoZona,
        string? codigoDistrito,
        string? tipoPedido = null)
    {
        if (string.IsNullOrWhiteSpace(codigoPedido))
            throw new DomainException("El código de pedido es obligatorio.", "PEDIDO_DELIVERY_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoTurno))
            throw new DomainException("El turno es obligatorio.", "PEDIDO_DELIVERY_TURNO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoClienteDelivery))
            throw new DomainException("El cliente delivery es obligatorio para un pedido delivery.", "PEDIDO_DELIVERY_CLIENTE_REQUERIDO");

        return new PedidoDelivery
        {
            CodigoPedido = codigoPedido,
            CodigoTurno = codigoTurno,
            CodigoUsuario = codigoUsuario,
            CodigoCaja = codigoCaja,
            CodigoClienteDelivery = codigoClienteDelivery,
            CodigoZona = codigoZona,
            CodigoDistrito = codigoDistrito,
            TipoPedido = tipoPedido,
            EstadoDelivery = EstadoDelivery.Pendiente,
            FechaRegistro = DateTime.Now
        };
    }

    /// <summary>
    /// Asigna motorizado al pedido delivery.
    /// Legacy: <c>frmAsignacionMotorizado.frm → Form_Load</c> — actualiza campo en MPEDIDO.
    /// Regla BR-DEL-001.
    /// </summary>
    public void AsignarMotorizado(string codigoMotorizado)
    {
        if (EstadoDelivery == EstadoDelivery.Entregado || EstadoDelivery == EstadoDelivery.Anulado)
            throw new DomainException("No se puede asignar motorizado a un pedido entregado o anulado.", "PEDIDO_DELIVERY_ESTADO_INVALIDO");

        CodigoMotorizado = codigoMotorizado;
        EstadoDelivery = EstadoDelivery.EnCamino;
        FechaDespacho = DateTime.Now;
    }

    /// <summary>
    /// Registra la entrega del pedido delivery.
    /// Legacy: <c>fmrPedidosEntregado.frm</c> — acción de confirmar entrega.
    /// Regla BR-DEL-009.
    /// </summary>
    public void RegistrarEntrega()
    {
        if (EstadoDelivery != EstadoDelivery.EnCamino)
            throw new DomainException("Solo se puede registrar entrega si el pedido está en camino.", "PEDIDO_DELIVERY_NO_EN_CAMINO");

        EstadoDelivery = EstadoDelivery.Entregado;
        FechaEntrega = DateTime.Now;
    }

    public void ActualizarEstado(EstadoDelivery nuevoEstado)
    {
        if (EstadoDelivery == EstadoDelivery.Anulado)
            throw new DomainException("No se puede cambiar el estado de un pedido anulado.", "PEDIDO_DELIVERY_ANULADO");

        EstadoDelivery = nuevoEstado;
    }

    public void EstablecerHoraEstimada(DateTime horaEstimada)
    {
        HoraEstimadaEntrega = horaEstimada;
    }
}
