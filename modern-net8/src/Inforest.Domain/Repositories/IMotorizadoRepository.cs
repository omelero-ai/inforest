using Inforest.Domain.Entities.Motorizado;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Repositorio de motorizados de delivery.
/// <para>
/// Legacy: tabla <c>TMOTORIZADODATOS</c>, vista <c>vMotorizado</c>.
/// Formularios: <c>frmMotorizado.frm</c>, <c>frmMotorizadoDetalle.frm</c>,
///   <c>frmAsignacionMotorizado.frm</c>, <c>frmReasignacionMotorizado.frm</c>,
///   <c>frmLlegadaSalida.frm</c>, <c>frmTarifaMotorizado.frm</c>.
/// </para>
/// Reglas: BR-DEL-002, BR-DEL-005, BR-DEL-006.
/// </summary>
public interface IMotorizadoRepository
{
    Task<Motorizado?> ObtenerPorCodigoAsync(string codigo, CancellationToken cancellationToken = default);
    Task<IEnumerable<Motorizado>> ObtenerTodosAsync(CancellationToken cancellationToken = default);
    Task InsertarAsync(Motorizado motorizado, CancellationToken cancellationToken = default);
    Task ActualizarTarifasAsync(string codigo, decimal tarifaLV, decimal tarifaSD, decimal tarifaES, string usuarioModificacion, CancellationToken cancellationToken = default);
    Task<AsignacionMotorizado?> ObtenerAsignacionActivaAsync(string codigoPedido, CancellationToken cancellationToken = default);
    Task RegistrarAsignacionAsync(AsignacionMotorizado asignacion, CancellationToken cancellationToken = default);
    Task RegistrarSalidaAsync(string codigoPedido, string codigoMotorizado, CancellationToken cancellationToken = default);
    Task RegistrarLlegadaAsync(string codigoPedido, string codigoMotorizado, CancellationToken cancellationToken = default);
}
