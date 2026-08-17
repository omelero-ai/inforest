namespace Inforest.Application.Delivery;

/// <summary>
/// Interfaz de repositorio de lectura para búsqueda y estadísticas de clientes delivery.
/// <para>
/// Legacy: <c>frmBusquedaDelivery.frm</c> — POS-FUNC-036.
/// Queries: <c>TDELIVERY LEFT JOIN vZona</c>, <c>vDelivery</c>,
/// <c>MPEDIDO/DPEDIDO</c>, <c>MDOCUMENTO/vCliente</c>, <c>vTienda</c>.
/// </para>
/// </summary>
public interface IClienteDeliveryReadRepository
{
    /// <summary>
    /// Devuelve el listado principal de mantenimiento de clientes delivery.
    /// Legacy: <c>frmClienteDelivery.frm</c> Form_Load —
    /// <c>SELECT … FROM vDelivery</c>.
    /// Regla BR-DEL-037.
    /// </summary>
    Task<IReadOnlyList<ClienteDeliveryListadoItem>> ListarMantenimientoAsync(CancellationToken ct = default);

    /// <summary>
    /// Devuelve todos los clientes delivery activos con descripción de zona para la grilla de búsqueda.
    /// Legacy: Form_Load de <c>frmBusquedaDelivery.frm</c> —
    /// <c>SELECT … FROM TDELIVERY LEFT JOIN vZona … WHERE lActivo=1</c>.
    /// Regla BR-DEL-036.
    /// </summary>
    Task<IReadOnlyList<ClienteDeliveryBusquedaItem>> ListarActivosConZonaAsync(CancellationToken ct = default);

    /// <summary>
    /// Obtiene el detalle completo de un cliente delivery para el panel lateral.
    /// Legacy: Sub Asigna() de <c>frmBusquedaDelivery.frm</c> —
    /// <c>SELECT * FROM vDelivery WHERE Codigo=@codigo</c>.
    /// Regla BR-DEL-036.
    /// </summary>
    Task<ClienteDeliveryDetalleBusqueda?> ObtenerDetalleAsync(string codigoDelivery, CancellationToken ct = default);

    /// <summary>
    /// Calcula estadísticas históricas de pedidos y último documento del cliente.
    /// Legacy: cmdOpcion(3) "Otros Datos" de <c>frmBusquedaDelivery.frm</c>.
    /// Regla BR-DEL-036.
    /// </summary>
    Task<EstadisticasClienteDelivery> ObtenerEstadisticasAsync(string codigoDelivery, int diasHistorico, CancellationToken ct = default);

    /// <summary>
    /// Lista las tiendas/sucursales activas de un cliente delivery.
    /// Legacy: Tienda_Click de <c>frmBusquedaDelivery.frm</c> —
    /// <c>SELECT * FROM vTienda WHERE lActivo=1 AND tCodigoDelivery=@cod</c>.
    /// Regla BR-DEL-036.
    /// </summary>
    Task<IReadOnlyList<TiendaDeliveryItem>> ObtenerTiendasAsync(string codigoDelivery, CancellationToken ct = default);
}
