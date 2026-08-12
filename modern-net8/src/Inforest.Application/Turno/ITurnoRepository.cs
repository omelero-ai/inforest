using TurnoCaja = Inforest.Domain.Entities.Configuracion.Turno;

namespace Inforest.Application.Turno;

public interface ITurnoRepository
{
    Task<TurnoCaja?> ObtenerTurnoActualAsync(string codigoCaja, CancellationToken ct = default);
    Task<bool> InsertarAsync(TurnoCaja turno, CancellationToken ct = default);
    Task<bool> CerrarAsync(string codigoTurno, decimal montoFinal, CancellationToken ct = default);
    Task<IReadOnlyList<TurnoCaja>> ObtenerHistorialAsync(string codigoCaja, DateTime desde, CancellationToken ct = default);
}
