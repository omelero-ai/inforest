namespace Inforest.Domain.Entities.Delivery;

/// <summary>
/// Fila del panel de despacho de delivery.
/// Legacy: <c>frmDespachador.frm</c> + vista <c>vDespachador</c>.
/// </summary>
public sealed record PedidoDespachadorResumen(
    string CodigoPedido,
    DateTime FechaRegistro,
    string? Usuario,
    string? Caja,
    string? Cliente,
    string? Telefono,
    string? Direccion,
    string? Referencia,
    string? Empacador,
    string? Motorizado,
    DateTime? FechaAsignacion,
    decimal MontoTotal);

/// <summary>
/// Opción de catálogo para selección de operador en despacho.
/// Legacy: vistas <c>vMotorizado</c> y <c>vEmpacador</c>.
/// </summary>
public sealed record OperadorDespachoItem(string Codigo, string Descripcion);
