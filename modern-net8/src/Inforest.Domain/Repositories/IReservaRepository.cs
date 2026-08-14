using Inforest.Domain.Entities.Reservas;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Repositorio de reservas — tabla TRESERVA.
/// Legacy: frmReserva.frm, frmReservaDetalle.frm
/// </summary>
public interface IReservaRepository
{
    Task<Reserva?> ObtenerPorCodigoAsync(string codigoReserva, CancellationToken cancellationToken = default);

    Task<IEnumerable<Reserva>> ObtenerPorFechaAsync(DateTime fechaInicio, DateTime fechaFin, CancellationToken cancellationToken = default);

    Task<string> GenerarProximoCodigoAsync(CancellationToken cancellationToken = default);

    Task InsertarAsync(Reserva reserva, CancellationToken cancellationToken = default);

    Task ActualizarAsync(Reserva reserva, CancellationToken cancellationToken = default);

    /// <summary>
    /// Convierte una reserva en pedido activo via spIns_MPEDIDO_RESERVA.
    /// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 4
    /// Retorna el código del pedido generado, o vacío si falló.
    /// </summary>
    Task<string> ConvertirAPedidoAsync(
        string codigoReserva,
        string codigoCaja,
        string codigoUsuario,
        string codigoTurno,
        DateTime fechaDiaContable,
        CancellationToken cancellationToken = default);
}
