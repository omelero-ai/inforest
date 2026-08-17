using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Repositorio de motivos de anulación.
/// Legacy: vMotivoAnulacion / TTABLA WHERE TTABLA = 'MOTIVOANULACION'.
/// </summary>
public interface IMotivoAnulacionRepository
{
    /// <summary>
    /// Obtiene los motivos de anulación activos, ordenados por código.
    /// BR-MOTIVO-001.
    /// </summary>
    Task<IReadOnlyList<MotivoAnulacion>> ObtenerActivosAsync(CancellationToken ct = default);
}
