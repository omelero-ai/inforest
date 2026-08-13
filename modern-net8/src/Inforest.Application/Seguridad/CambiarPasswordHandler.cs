using Inforest.Application.Interfaces;
using Inforest.Domain.Common;

namespace Inforest.Application.Seguridad;

// ─────────────────────────────────────────────────────────────────────────────
// Command
// ─────────────────────────────────────────────────────────────────────────────

public sealed record CambiarPasswordCommand(
    string LoginUsuario,
    string PasswordActual,
    string PasswordNuevo,
    string PasswordNuevoConfirmacion);

// ─────────────────────────────────────────────────────────────────────────────
// Handler
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Cambia la contraseña de un usuario POS.
/// Legacy: frmCambiarContrasenia.frm CmdAceptar_Click. BR-006.
/// Delega al servicio de autenticación que tiene acceso al cifrado y la BD.
/// </summary>
public sealed class CambiarPasswordHandler
{
    private readonly IAuthService _authService;

    public CambiarPasswordHandler(IAuthService authService)
        => _authService = authService;

    public async Task<Result> HandleAsync(CambiarPasswordCommand command, CancellationToken ct = default)
    {
        // ── Validaciones de aplicación ────────────────────────────────────────
        if (string.IsNullOrWhiteSpace(command.LoginUsuario))
            return Result.Fail("Ingrese el usuario.", "SEGURIDAD_USUARIO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(command.PasswordActual))
            return Result.Fail("Ingrese la contraseña actual.", "SEGURIDAD_PASSWORD_ACTUAL_REQUERIDO");

        if (string.IsNullOrWhiteSpace(command.PasswordNuevo))
            return Result.Fail("Ingrese la nueva contraseña.", "SEGURIDAD_PASSWORD_NUEVO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(command.PasswordNuevoConfirmacion))
            return Result.Fail("Confirme la nueva contraseña.", "SEGURIDAD_PASSWORD_CONFIRMACION_REQUERIDA");

        // Legacy: "La nueva contraseña debe ser distinta a la actual"
        if (string.Equals(command.PasswordActual.Trim(), command.PasswordNuevo.Trim(), StringComparison.Ordinal))
            return Result.Fail("La nueva contraseña debe ser distinta a la actual.", "SEGURIDAD_PASSWORD_IGUAL_A_ACTUAL");

        // Legacy: "Nueva contraseña no coincide con la confirmación"
        if (!string.Equals(command.PasswordNuevo.Trim(), command.PasswordNuevoConfirmacion.Trim(), StringComparison.Ordinal))
            return Result.Fail("Nueva contraseña no coincide con la confirmación.", "SEGURIDAD_PASSWORD_CONFIRMACION_NO_COINCIDE");

        return await _authService.CambiarPasswordAsync(
            command.LoginUsuario.Trim(),
            command.PasswordActual.Trim(),
            command.PasswordNuevo.Trim(),
            ct);
    }
}
