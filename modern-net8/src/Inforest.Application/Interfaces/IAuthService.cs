namespace Inforest.Application.Interfaces;

/// <summary>
/// Abstracción para el servicio de seguridad/autenticación.
/// Legacy: ClsSeguridad.cls — cifrado XOR+César (REEMPLAZAR con BCrypt).
/// Tablas: TUSUARIO, TGRUPOUSUARIO, TACCESO, TGRUPOACCESO en INFSEGURIDAD.
/// </summary>
public interface IAuthService
{
    /// <summary>
    /// Autentica un usuario con contraseña.
    /// Legacy: ClsSeguridad.VerificarPassword() + TUSUARIO
    /// IMPORTANTE: La contraseña legacy usa XOR+César — migrar gradualmente con rehash en login.
    /// </summary>
    Task<AuthResult> AutenticarAsync(string codigoUsuario, string contrasena, CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifica si un usuario tiene acceso a una funcionalidad.
    /// Legacy: TACCESO/TGRUPOACCESO
    /// </summary>
    Task<bool> TieneAccesoAsync(string codigoUsuario, string codigoAcceso, CancellationToken cancellationToken = default);
}

public record AuthResult(bool Exitoso, string? CodigoUsuario, string? NombreUsuario, string? MensajeError);
