using Inforest.Domain.Entities.Caja;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Contrato de acceso a datos para recibos de ingreso de caja.
/// Legacy: MINGRESO / vingreso. BR-RECIBO-001, BR-RECIBO-006.
/// </summary>
public interface IReciboIngresoRepository
{
    /// <summary>
    /// Obtiene recibos de ingreso en un rango de fechas.
    /// BR-RECIBO-006: Si filtrarPorUsuario=true filtra por usuario (lMCPV).
    /// </summary>
    Task<IReadOnlyList<ReciboIngreso>> ObtenerPorFechaAsync(
        DateTime desde,
        DateTime hasta,
        string? codigoUsuario = null,
        CancellationToken ct = default);

    Task<ReciboIngreso?> ObtenerPorCodigoAsync(string codigoRecibo, CancellationToken ct = default);

    /// <summary>
    /// Inserta un nuevo recibo de ingreso. BR-RECIBO-001: Código YY+8 ya calculado.
    /// </summary>
    Task InsertarAsync(ReciboIngreso recibo, CancellationToken ct = default);

    /// <summary>
    /// Actualiza el estado del recibo (ej. anulación). BR-RECIBO-003.
    /// </summary>
    Task ActualizarEstadoAsync(string codigoRecibo, string nuevoEstado, CancellationToken ct = default);

    /// <summary>
    /// Obtiene el máximo correlativo del año actual para generar el próximo código.
    /// BR-RECIBO-001: SELECT MAX(tRecibo) FROM MINGRESO WHERE SUBSTRING(tRecibo,1,2)=YY
    /// </summary>
    Task<string?> ObtenerMaximoCorrelativoAnioAsync(string anoCorto, CancellationToken ct = default);
}
