using Inforest.Domain.Entities.Delivery;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Repositorio de pedidos delivery.
/// <para>
/// Legacy: tablas <c>MPEDIDO</c>/<c>DPEDIDO</c>, vistas <c>vDespachador</c>,
/// <c>vDelivery</c>, <c>vPedidoCabecera</c>, <c>vPedidoDetalle</c>, <c>vPedidoGrilla</c>.
/// Formularios: <c>frmDespachador.frm</c>, <c>frmPedidoDelivery.frm</c>, <c>frmCentralPedidos.frm</c>.
/// </para>
/// </summary>
public interface IPedidoDeliveryRepository
{
    Task<PedidoDelivery?> ObtenerPorCodigoAsync(string codigoPedido, CancellationToken cancellationToken = default);
    Task<IEnumerable<PedidoDelivery>> ObtenerPendientesAsync(string codigoCaja, DateTime fechaInicio, DateTime fechaFin, CancellationToken cancellationToken = default);
    Task<IEnumerable<PedidoDelivery>> ObtenerPorEstadoAsync(EstadoDelivery estado, DateTime fechaInicio, DateTime fechaFin, CancellationToken cancellationToken = default);
    Task<IEnumerable<PedidoReasignacionMotorizado>> ObtenerPedidosReasignacionAsync(DateTime fechaInicio, DateTime fechaFin, string? criterioPedido, CancellationToken cancellationToken = default);
    Task<IEnumerable<PedidoDelivery>> ObtenerParaDespachadorAsync(DateTime fechaInicio, DateTime fechaFin, CancellationToken cancellationToken = default);
    Task InsertarAsync(PedidoDelivery pedido, CancellationToken cancellationToken = default);
    Task ActualizarEstadoAsync(string codigoPedido, EstadoDelivery estado, CancellationToken cancellationToken = default);
    Task AsignarMotorizadoAsync(string codigoPedido, string codigoMotorizado, CancellationToken cancellationToken = default);
    Task<int> ContarAsignacionesPrincipalesMotorizadoAsync(string codigoMotorizado, DateTime fechaOperacion, CancellationToken cancellationToken = default);
    Task<int> ContarReasignacionesMotorizadoAsync(string codigoMotorizado, DateTime fechaOperacion, CancellationToken cancellationToken = default);
    Task ActualizarReasignacionMotorizadoAsync(string codigoPedido, string codigoMotorizado, decimal tarifaMotorizado, bool esTarifaExtra, CancellationToken cancellationToken = default);
    Task LimpiarReasignacionMotorizadoAsync(string codigoPedido, CancellationToken cancellationToken = default);

    // ── CentralPedidos / frmCentralPedidos.frm ────────────────────────────────

    /// <summary>
    /// Confirma la entrega de un pedido (lEntregado = 1).
    /// Legacy: <c>frmCentralPedidos.frm Case 3</c>.
    /// SQL: <c>Update MPEDIDO Set lEntregado='1', tusuarioentregado=@usuario, fregentregado=getdate()</c>.
    /// Regla BR-DEL-012.
    /// </summary>
    Task ConfirmarEntregaAsync(string codigoPedido, string usuario, CancellationToken cancellationToken = default);

    /// <summary>
    /// Revierte la confirmación de entrega (lEntregado = 0).
    /// Legacy: <c>frmCentralPedidos.frm Case 5</c>.
    /// Regla BR-DEL-012.
    /// </summary>
    Task RevertirEntregaAsync(string codigoPedido, string usuario, CancellationToken cancellationToken = default);

    /// <summary>
    /// Modifica la fecha programada de entrega del pedido.
    /// Legacy: <c>frmCentralPedidos.frm Case 2</c>.
    /// SQL: <c>Update MPEDIDO set fregistro=..., fProgramacion=...</c>.
    /// Regla BR-DEL-014.
    /// </summary>
    Task ModificarFechaProgramadaAsync(string codigoPedido, DateTime nuevaFecha, CancellationToken cancellationToken = default);

    /// <summary>
    /// Obtiene el estado de pago del pedido desde <c>vDocumentoAgrupado</c>.
    /// Legacy: <c>frmCentralPedidos.frm Case 3</c> — comprueba si <c>Estado = 'POR COBRAR'</c>.
    /// Regla BR-DEL-013.
    /// </summary>
    Task<string?> ObtenerEstadoPagoAsync(string codigoPedido, CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifica si el pedido ya está marcado como entregado (lEntregado = 1).
    /// Legacy: <c>frmCentralPedidos.frm</c> — checks <c>grdGrilla.Columns(6).Text = "ENTREGADO"</c>.
    /// </summary>
    Task<bool> EstaEntregadoAsync(string codigoPedido, CancellationToken cancellationToken = default);
}
