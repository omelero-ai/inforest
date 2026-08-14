using BCrypt.Net;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Seguridad;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Security;

/// <summary>
/// Autenticación con compatibilidad Legacy + transición a BCrypt.
/// Legacy: frmAcceso.frm + usp_Inforest_ObtieneUsuarios + ClsSeguridad.cls.
/// </summary>
internal sealed class AuthService : IAuthService
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;
    private readonly IRbacService _rbacService;
    private readonly ISessionService _sessionService;
    private readonly IAuditoriaService _auditoriaService;
    private readonly IModernPasswordHashStore _modernPasswordHashStore;
    private readonly ILogger<AuthService> _logger;

    public AuthService(
        IDbConnectionFactory connectionFactory,
        ISpExecutor spExecutor,
        IRbacService rbacService,
        ISessionService sessionService,
        IAuditoriaService auditoriaService,
        IModernPasswordHashStore modernPasswordHashStore,
        ILogger<AuthService> logger)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
        _rbacService = rbacService;
        _sessionService = sessionService;
        _auditoriaService = auditoriaService;
        _modernPasswordHashStore = modernPasswordHashStore;
        _logger = logger;
    }

    public async Task<AuthResult> AutenticarAsync(AuthRequest request, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(request.CodigoUsuario))
            return AuthResult.Fallido("Ingrese su usuario.", "SEGURIDAD_USUARIO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(request.Contrasena) && string.IsNullOrWhiteSpace(request.BandaMagnetica))
            return AuthResult.Fallido("Ingrese su contraseña.", "SEGURIDAD_PASSWORD_REQUERIDO");

        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var records = await _spExecutor.QueryAsync<LegacyModuleUserRecord>(
            connection,
            "usp_Inforest_ObtieneUsuarios",
            new { Modulo = request.Modulo.Trim().ToUpperInvariant() },
            cancellationToken: cancellationToken);

        var user = records.FirstOrDefault(record =>
            string.Equals(record.tResumido, request.CodigoUsuario.Trim(), StringComparison.OrdinalIgnoreCase));

        if (user is null)
            return AuthResult.Fallido("Usuario no encontrado para el módulo solicitado.", "SEGURIDAD_USUARIO_NO_ENCONTRADO");

        if (!user.lActivo)
            return AuthResult.Fallido("El usuario se encuentra inactivo.", "SEGURIDAD_USUARIO_INACTIVO");

        var verification = await VerifyPasswordAsync(user, request, cancellationToken);
        if (!verification.Verified)
            return AuthResult.Fallido("Password erróneo.", "SEGURIDAD_PASSWORD_INVALIDO");

        var moduloSeguridad = ResolveModuloCode(request.Modulo);
        var permissions = await _rbacService.ObtenerPermisosAsync(user.tResumido, moduloSeguridad, cancellationToken);
        var auditResult = await _auditoriaService.RegistrarIngresoAsync(
            new RegistroAccesoAuditoriaRequest("I", request.BaseDatos, moduloSeguridad, user.tResumido.ToUpperInvariant(), 0),
            cancellationToken);

        if (!auditResult.EsExitoso)
            return AuthResult.Fallido(auditResult.MensajeError ?? "No se pudo registrar la auditoría de ingreso.", auditResult.CodigoError ?? "SEGURIDAD_AUDITORIA_INGRESO");

        var session = SesionOperativa.Iniciar(
            user.tCodigoUsuario,
            user.tDetallado,
            user.tGrupoUsuario,
            moduloSeguridad,
            request.CodigoCaja,
            request.CodigoTerminal,
            request.BaseDatos,
            auditResult.Valor,
            DateTime.UtcNow,
            permissions);

        var sessionResult = _sessionService.EstablecerSesion(session);
        if (!sessionResult.EsExitoso)
            return AuthResult.Fallido(sessionResult.MensajeError ?? "No se pudo abrir la sesión.", sessionResult.CodigoError ?? "SEGURIDAD_SESION_ERROR");

        return AuthResult.ExitosoConSesion(session, verification.PasswordMigrated, verification.PasswordUpgradePending);
    }

    public Task<bool> TieneAccesoAsync(string codigoUsuario, string modulo, string codigoAcceso, CancellationToken cancellationToken = default)
        => _rbacService.TieneAccesoAsync(codigoUsuario, modulo, codigoAcceso, cancellationToken);

    public async Task<Result> CerrarSesionAsync(CancellationToken cancellationToken = default)
    {
        var session = _sessionService.SesionActual;
        if (session is null)
            return Result.Fail("No existe una sesión activa.", "SEGURIDAD_SESION_INEXISTENTE");

        var auditResult = await _auditoriaService.RegistrarSalidaAsync(
            new RegistroAccesoAuditoriaRequest("S", session.BaseDatos, session.Modulo, session.CodigoUsuario, session.CorrelativoAcceso),
            cancellationToken);

        if (!auditResult.EsExitoso)
            return auditResult;

        return _sessionService.CerrarSesion(DateTime.UtcNow);
    }

    public async Task<Result> CambiarPasswordAsync(
        string loginUsuario,
        string passwordActual,
        string passwordNuevo,
        CancellationToken cancellationToken = default)
    {
        try
        {
            using var connection = await _connectionFactory.CreateOpenConnectionAsync("Infseguridad", cancellationToken);

            // Legacy: SELECT count(*) FROM tusuario WHERE tResumido=@Login AND lActivo=1
            var existe = await connection.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM TUSUARIO WHERE tResumido = @Login AND lActivo = 1",
                new { Login = loginUsuario }) > 0;

            if (!existe)
                return Result.Fail("El usuario no está activo o no está registrado.", "SEGURIDAD_USUARIO_INACTIVO");

            // Obtener código y hash almacenado
            var codigoUsuario = await connection.ExecuteScalarAsync<string?>(
                "SELECT tCodigoUsuario FROM TUSUARIO WHERE tResumido = @Login AND lActivo = 1",
                new { Login = loginUsuario });

            if (codigoUsuario is null)
                return Result.Fail("No se pudo localizar el usuario.", "SEGURIDAD_USUARIO_NO_ENCONTRADO");

            // Verificar contraseña actual (BCrypt primero, Legacy fallback)
            var verificado = false;
            var modernHash = await _modernPasswordHashStore.GetHashAsync(codigoUsuario, cancellationToken);
            if (!string.IsNullOrWhiteSpace(modernHash))
                verificado = BCrypt.Net.BCrypt.Verify(passwordActual, modernHash);

            if (!verificado)
            {
                var legacyHash = await connection.ExecuteScalarAsync<string?>(
                    "SELECT tPassword FROM TUSUARIO WHERE tResumido = @Login AND lActivo = 1",
                    new { Login = loginUsuario });

                if (!string.IsNullOrWhiteSpace(legacyHash))
                {
                    if (LooksLikeBcrypt(legacyHash))
                        verificado = BCrypt.Net.BCrypt.Verify(passwordActual, legacyHash);
                    else
                        verificado = string.Equals(
                            LegacyPasswordCipher.Decrypt(legacyHash).Trim(),
                            passwordActual.Trim().ToUpperInvariant(),
                            StringComparison.OrdinalIgnoreCase);
                }
            }

            if (!verificado)
                return Result.Fail("Contraseña actual incorrecta.", "SEGURIDAD_PASSWORD_ACTUAL_INVALIDO");

            // Legacy: "Clave no permitida" — verifica que la nueva no exista ya en la tabla
            var encriptadoNuevo = LegacyPasswordCipher.Encrypt(passwordNuevo);
            var claveEnUso = await connection.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM TUSUARIO WHERE tPassword = @Password AND tResumido <> @Login",
                new { Password = encriptadoNuevo, Login = loginUsuario }) > 0;

            if (claveEnUso)
                return Result.Fail("Clave no permitida, intente de nuevo.", "SEGURIDAD_PASSWORD_NO_PERMITIDO");

            // Actualizar contraseña Legacy en TUSUARIO
            await connection.ExecuteAsync(
                """
                UPDATE TUSUARIO
                   SET tPassword        = @Password,
                       fRegistro        = GETDATE(),
                       tUsuarioModifica = @UsuarioModifica
                 WHERE tCodigoUsuario = @CodigoUsuario
                   AND lActivo = 1
                """,
                new { Password = encriptadoNuevo, UsuarioModifica = loginUsuario, CodigoUsuario = codigoUsuario });

            // Actualizar hash BCrypt moderno en tabla sidecar TUSUARIO_HASH (ADR-013, SEC-006)
            await _modernPasswordHashStore.UpsertHashAsync(
                codigoUsuario,
                BCrypt.Net.BCrypt.HashPassword(passwordNuevo),
                cancellationToken);

            _logger.LogInformation("Contraseña cambiada para usuario {Login}.", loginUsuario);
            return Result.Ok();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al cambiar contraseña del usuario {Login}.", loginUsuario);
            return Result.Fail("Error al cambiar la contraseña. Intente nuevamente.", "SEGURIDAD_ERROR_INTERNO");
        }
    }

    internal static string ResolveModuloCode(string modulo)
        => modulo.Trim().ToUpperInvariant() switch
        {
            "INFOREST" => "01",
            "ADMINISTRACION" => "02",
            _ => "03"
        };

    private async Task<PasswordVerificationResult> VerifyPasswordAsync(LegacyModuleUserRecord user, AuthRequest request, CancellationToken cancellationToken)
    {
        var normalizedPassword = request.Contrasena.Trim().ToUpperInvariant();
        var modernHash = await _modernPasswordHashStore.GetHashAsync(user.tCodigoUsuario, cancellationToken);

        if (!string.IsNullOrWhiteSpace(modernHash) && BCrypt.Net.BCrypt.Verify(request.Contrasena, modernHash))
            return PasswordVerificationResult.Success(passwordMigrated: false, passwordUpgradePending: false);

        if (LooksLikeBcrypt(user.tPassword) && BCrypt.Net.BCrypt.Verify(request.Contrasena, user.tPassword))
            return PasswordVerificationResult.Success(passwordMigrated: false, passwordUpgradePending: false);

        if (!string.IsNullOrWhiteSpace(user.tPassword) &&
            string.Equals(LegacyPasswordCipher.Decrypt(user.tPassword).Trim(), normalizedPassword, StringComparison.OrdinalIgnoreCase))
        {
            return await UpgradeLegacyPasswordAsync(user.tCodigoUsuario, request.Contrasena, cancellationToken);
        }

        if (!string.IsNullOrWhiteSpace(request.BandaMagnetica) &&
            !string.IsNullOrWhiteSpace(user.tBandaMagnetica) &&
            string.Equals(LegacyPasswordCipher.Decrypt(user.tBandaMagnetica).Trim(), request.BandaMagnetica.Trim().ToUpperInvariant(), StringComparison.OrdinalIgnoreCase))
        {
            return PasswordVerificationResult.Success(passwordMigrated: false, passwordUpgradePending: false);
        }

        return PasswordVerificationResult.Failure();
    }

    private async Task<PasswordVerificationResult> UpgradeLegacyPasswordAsync(string codigoUsuario, string plainPassword, CancellationToken cancellationToken)
    {
        var migrated = await _modernPasswordHashStore.UpsertHashAsync(
            codigoUsuario,
            BCrypt.Net.BCrypt.HashPassword(plainPassword),
            cancellationToken);

        if (!migrated)
            _logger.LogWarning("No se pudo persistir hash BCrypt para {CodigoUsuario}. Se seguirá usando compatibilidad legacy.", codigoUsuario);

        return PasswordVerificationResult.Success(passwordMigrated: migrated, passwordUpgradePending: !migrated);
    }

    private static bool LooksLikeBcrypt(string? value)
        => !string.IsNullOrWhiteSpace(value) && value.StartsWith("$2", StringComparison.Ordinal);

    private readonly record struct PasswordVerificationResult(bool Verified, bool PasswordMigrated, bool PasswordUpgradePending)
    {
        public static PasswordVerificationResult Success(bool passwordMigrated, bool passwordUpgradePending)
            => new(true, passwordMigrated, passwordUpgradePending);

        public static PasswordVerificationResult Failure()
            => new(false, false, false);
    }
}
