namespace Inforest.Domain.Entities.Delivery;

/// <summary>
/// Fila de seguimiento de pedidos delivery para paneles de monitoreo.
/// Legacy: <c>frmPedidoDelivery.frm</c>, <c>frmPedidoDeliveryNo.frm</c>, vista <c>vDespachador</c>.
/// </summary>
public sealed record PedidoDeliverySeguimiento(
    string CodigoPedido,
    DateTime FechaRegistro,
    string? Usuario,
    string? Caja,
    string? Telefono,
    string? Cliente,
    string? Empacador,
    string? Motorizado,
    DateTime? FechaAsignacion,
    DateTime? FechaSalida,
    DateTime? FechaLlegada,
    string? Referencia,
    string? Direccion,
    string? Zona);
