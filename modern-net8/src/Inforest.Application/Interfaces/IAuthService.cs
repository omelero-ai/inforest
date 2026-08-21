using Inforest.Domain.Common;
using Inforest.Domain.Entities.Seguridad;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Abstracción para el servicio de seguridad/autenticación.
/// Legacy: frmAcceso.frm + modProcedimiento.bas (Encapsula/Desencapsula) + transición BCrypt.
/// Tablas: TUSUARIO, TGRUPOUSUARIO, TACCESO, TGRUPOACCESO en INFSEGURIDAD.
/// </summary>
public interface IAuthService
{
    /// <summary>
    /// Autentica un usuario con contraseña/sesión base.
    /// Legacy: frmAcceso.frm + usp_Inforest_ObtieneUsuarios + Desencapsula().
    /// </summary>
    Task<AuthResult> AutenticarAsync(AuthRequest request, CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifica si un usuario tiene acceso a una funcionalidad.
    /// Legacy: TACCESO/TGRUPOACCESO.
    /// </summary>
    Task<bool> TieneAccesoAsync(string codigoUsuario, string modulo, string codigoAcceso, CancellationToken cancellationToken = default);

    /// <summary>
    /// Cierra la sesión actual y registra la salida de auditoría.
    /// </summary>
    Task<Result> CerrarSesionAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Cambia la contraseña de un usuario validando la actual.
    /// Legacy: frmCambiarContrasenia.frm CmdAceptar_Click. BR-006.
    /// Tablas: TUSUARIO + TUSUARIO_HASH (ADR-013, SEC-006).
    /// </summary>
    Task<Result> CambiarPasswordAsync(
        string loginUsuario,
        string passwordActual,
        string passwordNuevo,
        CancellationToken cancellationToken = default);
}

public sealed record AuthRequest(
    string CodigoUsuario,
    string Contrasena,
    string Modulo,
    string CodigoCaja,
    string CodigoTerminal,
    string BaseDatos,
    string? BandaMagnetica = null);

public sealed record AuthResult(
    bool Exitoso,
    SesionOperativa? Sesion,
    string? CodigoError,
    string? MensajeError,
    bool PasswordMigrated,
    bool PasswordUpgradePending)
{
    public static AuthResult Fallido(string mensaje, string codigoError)
       => new(false, null, codigoError, mensaje, false, false);

    public static AuthResult ExitosoConSesion(SesionOperativa sesion, bool passwordMigrated, bool passwordUpgradePending)
       => new(true, sesion, null, null, passwordMigrated, passwordUpgradePending);
}
