using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Ventas;

/// <summary>
/// Repositorio de correlativos de documento.
/// Legacy: TTIPODOCUMENTOIMPRESORA. BR-POS-008.
/// </summary>
public interface ICorrelativoRepository
{
    /// <summary>
    /// Obtiene todos los correlativos de una caja.
    /// Legacy: SELECT * FROM TTIPODOCUMENTOIMPRESORA WHERE tCaja = @tCaja
    /// </summary>
    Task<IReadOnlyList<CorrelativoDocumento>> ObtenerPorCajaAsync(string codigoCaja, CancellationToken ct = default);

    /// <summary>
    /// Obtiene el correlativo de FACTURA de una caja específica para la pantalla de inicio (BR-POS-008).
    /// Legacy: SELECT * FROM TTIPODOCUMENTOIMPRESORA WHERE tCaja=@sCaja AND tTipoEmision=@tipoEmision
    /// </summary>
    Task<CorrelativoDocumento?> ObtenerPorCajaYTipoAsync(string codigoCaja, string tipoEmision, CancellationToken ct = default);

    /// <summary>
    /// Obtiene todos los correlativos para mostrarse en la pantalla frmDocumentoCorrelativo.
    /// Legacy: SELECT * FROM TTIPODOCUMENTOIMPRESORA (con JOIN a vistas para descripción)
    /// </summary>
    Task<IReadOnlyList<CorrelativoDocumento>> ObtenerTodosAsync(CancellationToken ct = default);
}
