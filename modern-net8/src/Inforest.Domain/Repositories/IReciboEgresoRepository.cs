using Inforest.Domain.Entities.Caja;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Contrato de acceso a datos para recibos de egreso de caja.
/// Legacy: MEGRESO / vegreso / spRep_ReciboEgreso. BR-RECIBO-007, BR-RECIBO-009.
/// </summary>
public interface IReciboEgresoRepository
{
    /// <summary>
    /// Obtiene recibos de egreso en un rango de fechas.
    /// BR-RECIBO-006: Si filtrarPorUsuario=true filtra por usuario (lMCPV).
    /// </summary>
    Task<IReadOnlyList<ReciboEgreso>> ObtenerPorFechaAsync(
        DateTime desde,
        DateTime hasta,
        string? codigoUsuario = null,
        CancellationToken ct = default);

    Task<ReciboEgreso?> ObtenerPorCodigoAsync(string codigoRecibo, CancellationToken ct = default);

    /// <summary>
    /// Inserta un nuevo recibo de egreso. BR-RECIBO-007: Código YY+8 ya calculado.
    /// </summary>
    Task InsertarAsync(ReciboEgreso recibo, CancellationToken ct = default);

    /// <summary>
    /// Actualiza el estado del recibo de egreso (ej. anulación). BR-RECIBO-008.
    /// </summary>
    Task ActualizarEstadoAsync(string codigoRecibo, string nuevoEstado, CancellationToken ct = default);

    /// <summary>
    /// Obtiene el máximo correlativo del año actual para generar el próximo código.
    /// BR-RECIBO-007: SELECT MAX(tRecibo) FROM MEGRESO WHERE SUBSTRING(tRecibo,1,2)=YY
    /// </summary>
    Task<string?> ObtenerMaximoCorrelativoAnioAsync(string anoCorto, CancellationToken ct = default);
}
