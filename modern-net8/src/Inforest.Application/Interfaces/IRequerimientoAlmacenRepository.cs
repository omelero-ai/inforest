using Inforest.Domain.Entities.Almacen;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Repositorio de requerimientos del módulo ALMACEN.
/// Legacy: vRequerimiento (ALMACEN DB), MREQUERIMIENTO, TRUTAAREA.
/// frmImportacionRequerimientos.frm, frmImportacionRequerimientoDetalle.frm.
/// BR-IMPORT-001, BR-IMPORT-002, BR-IMPORT-003, BR-IMPORT-004.
/// </summary>
public interface IRequerimientoAlmacenRepository
{
    /// <summary>
    /// Obtiene requerimientos pendientes de importar al POS dentro de un rango de fechas.
    /// Legacy: SELECT DISTINCT v.Rq, v.Fecha, v.Area, v.Responsable, v.Estado, ISNULL(v.tPedido,'')
    ///         FROM vRequerimiento v INNER JOIN TRUTAAREA T ON v.CodigoArea = T.tCodigoArea
    ///         WHERE v.CodEstado='02' AND T.lImportarPV=1 AND fechas...
    /// BR-IMPORT-001: Solo requerimientos en estado '02' con área habilitada (lImportarPV=1).
    /// </summary>
    Task<IReadOnlyList<RequerimientoAlmacen>> ObtenerPendientesAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default);

    /// <summary>
    /// Obtiene el detalle de artículos de un requerimiento con su enlace a INFOREST.
    /// Legacy: SELECT V.CodArt, V.Articulo, V.Cantidad, ISNULL(T.tCodigoProducto,'') As Enlace
    ///         FROM vRequerimiento V LEFT JOIN &lt;sMDB&gt;.dbo.TPRODUCTO T ON V.CodArt = T.tEnlace
    ///         WHERE V.Rq = ...
    /// BR-IMPORT-003: Muestra enlace para verificar que todos los artículos tienen mapeo.
    /// </summary>
    Task<IReadOnlyList<DetalleRequerimientoAlmacen>> ObtenerDetalleAsync(
        string rq,
        string baseDatosInforest,
        CancellationToken ct = default);

    /// <summary>
    /// Verifica si el requerimiento ya tiene un pedido generado.
    /// Legacy: SELECT ISNULL(tPedido,'') FROM MREQUERIMIENTO WHERE tRequerimiento = ...
    /// BR-IMPORT-002: No se puede importar un requerimiento ya procesado.
    /// </summary>
    Task<string?> ObtenerPedidoAsociadoAsync(
        string rq,
        CancellationToken ct = default);

    /// <summary>
    /// Marca el requerimiento como importado asociándolo al pedido generado.
    /// Legacy: UPDATE MREQUERIMIENTO SET lPedido=1, tPedido=@pedido WHERE tRequerimiento=@rq
    /// BR-IMPORT-004: Debe marcarse el requerimiento como procesado tras importar.
    /// </summary>
    Task MarcarImportadoAsync(
        string rq,
        string codigoPedido,
        CancellationToken ct = default);
}
