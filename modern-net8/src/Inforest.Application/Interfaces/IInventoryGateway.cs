namespace Inforest.Application.Interfaces;

/// <summary>
/// Contrato para descargo de almacén sobre ventas.
/// Legacy: clsAlmacen.cls — BR-008.
/// SPs: usp_Inforest_DescargoVenta, usp_Inforest_DescargoVentaPorInsumo, USP_MODIFICARINSUMOS.
/// </summary>
public interface IInventoryGateway
{
    /// <summary>
    /// Ejecuta el descargo de almacén para un pedido de venta.
    /// Legacy: usp_Inforest_DescargoVenta — llena tabla temporal y ejecuta descargo por receta/insumo.
    /// </summary>
    Task<bool> EjecutarDescargoVentaAsync(
        string codigoAlmacen,
        DateTime fechaInicio,
        DateTime fechaFin,
        string tablaTemporal,
        string codigoLocal,
        string codigoPedido,
        int tipoOperacion,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta el descargo por insumo para un pedido de venta.
    /// Legacy: usp_Inforest_DescargoVentaPorInsumo.
    /// </summary>
    Task<bool> EjecutarDescargoVentaPorInsumoAsync(
        string codigoPedido,
        string codigoAlmacen,
        DateTime fechaDiaContable,
        string usuario,
        CancellationToken ct = default);

    /// <summary>
    /// Modifica insumos vinculados a un pedido.
    /// Legacy: USP_MODIFICARINSUMOS.
    /// </summary>
    Task<bool> ModificarInsumosAsync(
        string codigoPedido,
        string codigoAlmacen,
        string usuario,
        CancellationToken ct = default);

    /// <summary>
    /// Registra el log de descargo.
    /// Legacy: usp_Inforest_InsertaLogDescargo.
    /// </summary>
    Task InsertarLogDescargoAsync(
        int correlativo,
        DateTime fechaRegistro,
        string usuario,
        DateTime fechaInicio,
        DateTime fechaFin,
        int cantidadPedidos,
        int cantidadPedidosOk,
        int cantidadPedidosError,
        string equipoIp,
        string equipoNombre,
        string equipoUsuario,
        CancellationToken ct = default);
}
