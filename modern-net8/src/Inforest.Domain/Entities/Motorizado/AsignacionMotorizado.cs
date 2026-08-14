using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Motorizado;

/// <summary>
/// Registro de asignación de un motorizado a un pedido delivery.
/// <para>
/// Legacy: relación entre <c>MPEDIDO.tMotorizado</c> y <c>TMOTORIZADODATOS</c>;
/// gestionada en <c>frmAsignacionMotorizado.frm</c> y <c>frmReasignacionMotorizado.frm</c>.
/// Tiempos de salida y llegada registrados en <c>frmLlegadaSalida.frm</c>.
/// </para>
/// Reglas: BR-DEL-001, BR-DEL-006.
/// </summary>
public class AsignacionMotorizado : Entity
{
    public string CodigoPedido { get; private set; } = string.Empty;
    public string CodigoMotorizado { get; private set; } = string.Empty;
    public DateTime FechaAsignacion { get; private set; }
    public DateTime? FechaSalida { get; private set; }
    public DateTime? FechaLlegada { get; private set; }
    public string UsuarioAsignacion { get; private set; } = string.Empty;
    public string? UsuarioReasignacion { get; private set; }
    public bool EsReasignacion { get; private set; }

    // Dapper constructor
    private AsignacionMotorizado() { }

    public static AsignacionMotorizado Crear(
        string codigoPedido,
        string codigoMotorizado,
        string usuarioAsignacion)
    {
        if (string.IsNullOrWhiteSpace(codigoPedido))
            throw new DomainException("El código de pedido es obligatorio.", "ASIGNACION_PEDIDO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoMotorizado))
            throw new DomainException("El código de motorizado es obligatorio.", "ASIGNACION_MOTORIZADO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(usuarioAsignacion))
            throw new DomainException("El usuario de asignación es obligatorio.", "ASIGNACION_USUARIO_REQUERIDO");

        return new AsignacionMotorizado
        {
            CodigoPedido = codigoPedido,
            CodigoMotorizado = codigoMotorizado,
            UsuarioAsignacion = usuarioAsignacion,
            FechaAsignacion = DateTime.Now,
            EsReasignacion = false
        };
    }

    /// <summary>
    /// Reasigna el pedido a otro motorizado.
    /// Legacy: <c>frmReasignacionMotorizado.frm</c>.
    /// </summary>
    public AsignacionMotorizado Reasignar(string nuevoCodigoMotorizado, string usuarioReasignacion)
    {
        if (string.IsNullOrWhiteSpace(nuevoCodigoMotorizado))
            throw new DomainException("El nuevo código de motorizado es obligatorio.", "REASIGNACION_MOTORIZADO_REQUERIDO");

        if (FechaLlegada.HasValue)
            throw new DomainException("No se puede reasignar un pedido ya entregado.", "REASIGNACION_YA_ENTREGADO");

        return new AsignacionMotorizado
        {
            CodigoPedido = CodigoPedido,
            CodigoMotorizado = nuevoCodigoMotorizado,
            UsuarioAsignacion = UsuarioAsignacion,
            UsuarioReasignacion = usuarioReasignacion,
            FechaAsignacion = DateTime.Now,
            EsReasignacion = true
        };
    }

    /// <summary>
    /// Registra salida del motorizado.
    /// Legacy: <c>frmLlegadaSalida.frm</c> — opción "Salida".
    /// Regla BR-DEL-006.
    /// </summary>
    public void RegistrarSalida()
    {
        if (FechaSalida.HasValue)
            throw new DomainException("La salida ya fue registrada.", "ASIGNACION_SALIDA_YA_REGISTRADA");

        FechaSalida = DateTime.Now;
    }

    /// <summary>
    /// Registra llegada (entrega) del motorizado.
    /// Legacy: <c>frmLlegadaSalida.frm</c> — opción "Llegada".
    /// Regla BR-DEL-006.
    /// </summary>
    public void RegistrarLlegada()
    {
        if (!FechaSalida.HasValue)
            throw new DomainException("No se puede registrar llegada sin haber registrado salida.", "ASIGNACION_SALIDA_PENDIENTE");

        if (FechaLlegada.HasValue)
            throw new DomainException("La llegada ya fue registrada.", "ASIGNACION_LLEGADA_YA_REGISTRADA");

        FechaLlegada = DateTime.Now;
    }
}
