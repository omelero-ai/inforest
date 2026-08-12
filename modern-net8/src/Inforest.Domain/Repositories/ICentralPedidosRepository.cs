using Inforest.Domain.Entities.Delivery;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Registro de pedido del Central de Pedidos multi-local.
/// <para>
/// Legacy: vista <c>vTablasCentralizada</c>, SP <c>sp_Inforest_PedidosCentralPedido</c>
/// en base <c>CENTRALDELIVERY</c>.
/// </para>
/// Regla BR-DEL-008.
/// </summary>
public sealed record PedidoCentral(
    string CodigoPedido,
    string? CodigoLocal,
    string? NombreLocal,
    DateTime FechaRegistro,
    string? EstadoPedido,
    string? ClienteDelivery,
    string? Telefono,
    string? Direccion,
    decimal MontoTotal);

/// <summary>
/// Repositorio para operaciones con la base <c>CENTRALDELIVERY</c>.
/// <para>
/// Legacy: conexión separada configurada en INI sección <c>[CentralDelivery]</c>.
/// Formularios: <c>frmCentralPedidos.frm</c>, <c>frmBusquedaCentralPedido.frm</c>.
/// SPs: <c>sp_Inforest_PedidosCentralPedido</c>, <c>sp_CD_Modificar_EstadoDelivery_Cabecera</c>.
/// </para>
/// Reglas: BR-DEL-003, BR-DEL-008, BR-DEL-009.
/// </summary>
public interface ICentralPedidosRepository
{
    /// <summary>
    /// Obtiene pedidos de Central de Pedidos.
    /// Legacy SP: <c>sp_Inforest_PedidosCentralPedido</c>.
    /// </summary>
    Task<IEnumerable<PedidoCentral>> ObtenerPedidosCentralAsync(DateTime fechaInicio, DateTime fechaFin, CancellationToken cancellationToken = default);

    /// <summary>
    /// Modifica el estado de un delivery en la base CENTRALDELIVERY.
    /// Legacy SP: <c>sp_CD_Modificar_EstadoDelivery_Cabecera</c>.
    /// Regla BR-DEL-009.
    /// </summary>
    Task ModificarEstadoDeliveryAsync(string codigoPedido, string nuevoEstado, string usuario, CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifica si la conexión a CENTRALDELIVERY está disponible.
    /// Legacy: <c>VerConexionInfoRestCentral()</c> en modDespachador.bas.
    /// Regla BR-DEL-003.
    /// </summary>
    Task<bool> VerificarConexionAsync(CancellationToken cancellationToken = default);
}
