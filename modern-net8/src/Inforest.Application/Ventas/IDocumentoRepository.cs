using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Ventas;

/// <summary>
/// DTO para ítems de DPEDIDO pendientes de facturación.
/// Legacy: SELECT from DPEDIDO where tFacturado='' AND tEstadoItem='N'. BR-DOC-001.
/// </summary>
public sealed record ItemPendienteFacturacionDto(
    string CodigoPedido,
    string Item,
    string CodigoProducto,
    string NombreProducto,
    decimal PrecioVenta,
    decimal Cantidad,
    decimal Venta,
    string? Area);

/// <summary>
/// DTO para documentos pendientes de cobro en una caja.
/// Legacy: SELECT from MDOCUMENTO where tEstadoDocumento='01'. BR-DOC-008.
/// </summary>
public sealed record DocumentoPendienteDto(
    string CodigoDocumento,
    string? Mesa,
    string? Motorizado,
    string? Observacion,
    string? CodigoPedido,
    decimal Total,
    DateTime FechaRegistro,
    string? Cliente,
    string TipoPedido,
    string TipoDocumento,
    string EstadoDocumento);

public interface IDocumentoRepository
{
    Task<Documento?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default);
    Task<bool> InsertarAsync(Documento doc, IReadOnlyList<DetalleDocumento> detalles, CancellationToken ct = default);
    Task<bool> MarcarCobradoAsync(string codigo, decimal abono, CancellationToken ct = default);
    Task<bool> AnularAsync(string codigo, string usuario, string? motivo = null, CancellationToken ct = default);

    /// <summary>
    /// Obtiene los ítems de DPEDIDO para un pedido que aún no han sido facturados.
    /// Legacy: SELECT DPEDIDO.*, TPRODUCTO.tResumido FROM DPEDIDO INNER JOIN TPRODUCTO
    /// WHERE tCodigoPedido=@pedido AND (ISNULL(tFacturado,'0')='0' OR LEN(LTRIM(tFacturado))=0)
    /// AND tEstadoItem='N'. BR-DOC-001.
    /// </summary>
    Task<IReadOnlyList<ItemPendienteFacturacionDto>> ObtenerItemsPendientesFacturacionAsync(
        string codigoPedido,
        CancellationToken ct = default);

    /// <summary>
    /// Obtiene los documentos pendientes de cobro (tEstadoDocumento='01') para una caja.
    /// Legacy: SELECT MDOCUMENTO.*, TCLIENTE.tEmpresa, TMESA.tResumido, MPEDIDO.tObservacion
    /// FROM MDOCUMENTO ... WHERE tEstadoDocumento='01' AND tCaja=@caja. BR-DOC-008.
    /// </summary>
    Task<IReadOnlyList<DocumentoPendienteDto>> ObtenerDocumentosPendientesCajaAsync(
        string codigoCaja,
        CancellationToken ct = default);

    /// <summary>
    /// Llama al SP usp_Inforest_Impresion para reimpresión de un documento.
    /// Legacy: EXEC usp_Inforest_Impresion @documento, @modo. BR-DOC-005.
    /// </summary>
    Task<bool> ReimprimirAsync(string codigoDocumento, int modo = 3, CancellationToken ct = default);
}
