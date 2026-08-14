using Inforest.Domain.Entities.Cocina;

namespace Inforest.Application.Kitchen;

/// <summary>
/// Contrato de mensajería operativa de cocina/KDS.
/// Legacy: frmMensajeCocina.frm, frmMensajeCocinaDetalle.frm, TMENSAJECOCINA, USP_*MENSAJE*.
/// Reglas BR-MSGCOC-001, BR-MSGCOC-002, BR-MSGCOC-003, BR-MSGCOC-004, BR-MSGCOC-005.
/// </summary>
public interface IMensajeCocinaRepository
{
    Task<IReadOnlyList<MensajeCocina>> ObtenerMensajesAsync(DateTime fechaInicio, DateTime fechaFin, string codigoCaja, CancellationToken ct = default);
    Task<MensajeCocina?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default);
    Task<int> ContarActivosAsync(string? codigoExcluir = null, CancellationToken ct = default);
    Task<string> ObtenerProximoCodigoAsync(DateTime fecha, CancellationToken ct = default);
    Task AgregarAsync(MensajeCocina mensaje, CancellationToken ct = default);
    Task ModificarAsync(MensajeCocina mensaje, CancellationToken ct = default);
    Task EliminarAsync(string codigo, CancellationToken ct = default);
    Task CerrarActivosPorCajaAsync(string usuario, string codigoCaja, CancellationToken ct = default);
}
