using Inforest.Domain.Entities.Caja;
using TurnoCaja = Inforest.Domain.Entities.Configuracion.Turno;

namespace Inforest.Application.Turno;

public interface ITurnoRepository
{
    Task<TurnoCaja?> ObtenerTurnoActualAsync(string codigoCaja, CancellationToken ct = default);
    Task<bool> InsertarAsync(TurnoCaja turno, CancellationToken ct = default);

    /// <summary>
    /// Cierra el turno actualizando MTURNO con lCierre=1 y el desglose completo de montos.
    /// Legacy: UPDATE MTURNO SET lCierre=1, fFinal=getdate(), nMontoEN=..., nMontoFN=..., etc.
    /// BR-CAJA-004.
    /// </summary>
    Task<bool> CerrarAsync(string codigoTurno, decimal montoFinal, CierreTurnoBreakdown breakdown, CancellationToken ct = default);

    Task<IReadOnlyList<TurnoCaja>> ObtenerHistorialAsync(string codigoCaja, DateTime desde, CancellationToken ct = default);
}
