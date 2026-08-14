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
}
