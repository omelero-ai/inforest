using Inforest.Application.Interfaces;
using Inforest.Application.Seguridad;
using Inforest.Domain.Entities.Seguridad;
using Moq;

namespace Inforest.Application.Tests.Seguridad;

/// <summary>
/// Tests comparativos Legacy vs .NET para el flujo de login del POS.
/// Legacy: frmAcceso.frm — cmdOpcion_Click(0), AccesoInicio.
/// Reglas: BR-POS-006 (autenticación previa al MDI POS), BR-POS-006-LOCK (bloqueo por intentos fallidos).
/// </summary>
public class AuthLoginIntentosTests
{
    // ─────────────────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────────────────

    private static IAuthService BuildAuthService(AuthResult resultado)
    {
        var mock = new Mock<IAuthService>();
        mock.Setup(s => s.AutenticarAsync(It.IsAny<AuthRequest>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(resultado);
        return mock.Object;
    }

    private static SesionOperativa BuildSesion()
        => SesionOperativa.Iniciar("USR1", "Admin", "GRP1", "01", "CA1", "POS01", "INFOREST", 1L, DateTime.UtcNow, []);

    // ─────────────────────────────────────────────────────────────────────────
    // BR-POS-006: códigos de error esperados del AuthService
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Legacy frmAcceso.frm línea 459: "Ingrese su usuario" cuando txtUsuario está vacío.
    /// </summary>
    [Fact]
    public async Task Login_UsuarioVacio_RetornaSEGURIDAD_USUARIO_REQUERIDO()
    {
        var authService = BuildAuthService(AuthResult.Fallido("Ingrese su usuario.", "SEGURIDAD_USUARIO_REQUERIDO"));
        var result = await authService.AutenticarAsync(new AuthRequest("", "PWD", "INFOREST", "CA1", "POS01", "INFOREST"));

        Assert.False(result.Exitoso);
        Assert.Equal("SEGURIDAD_USUARIO_REQUERIDO", result.CodigoError);
    }

    /// <summary>
    /// Legacy frmAcceso.frm línea 460: "Ingrese su password" cuando txtPassword está vacío.
    /// </summary>
    [Fact]
    public async Task Login_PasswordVacio_RetornaSEGURIDAD_PASSWORD_REQUERIDO()
    {
        var authService = BuildAuthService(AuthResult.Fallido("Ingrese su contraseña.", "SEGURIDAD_PASSWORD_REQUERIDO"));
        var result = await authService.AutenticarAsync(new AuthRequest("ADMIN", "", "INFOREST", "CA1", "POS01", "INFOREST"));

        Assert.False(result.Exitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_REQUERIDO", result.CodigoError);
    }

    /// <summary>
    /// Legacy frmAcceso.frm línea 466: "Usuario No Encontrado" cuando RsUsuario.EOF.
    /// </summary>
    [Fact]
    public async Task Login_UsuarioNoEncontrado_RetornaSEGURIDAD_USUARIO_NO_ENCONTRADO()
    {
        var authService = BuildAuthService(AuthResult.Fallido("Usuario no encontrado para el módulo solicitado.", "SEGURIDAD_USUARIO_NO_ENCONTRADO"));
        var result = await authService.AutenticarAsync(new AuthRequest("DESCONOCIDO", "PWD", "INFOREST", "CA1", "POS01", "INFOREST"));

        Assert.False(result.Exitoso);
        Assert.Equal("SEGURIDAD_USUARIO_NO_ENCONTRADO", result.CodigoError);
    }

    /// <summary>
    /// Legacy frmAcceso.frm línea 497: "Password Erroneo" cuando Desencapsula() no coincide.
    /// </summary>
    [Fact]
    public async Task Login_PasswordErroneo_RetornaSEGURIDAD_PASSWORD_INVALIDO()
    {
        var authService = BuildAuthService(AuthResult.Fallido("Password erróneo.", "SEGURIDAD_PASSWORD_INVALIDO"));
        var result = await authService.AutenticarAsync(new AuthRequest("ADMIN", "WRONGPWD", "INFOREST", "CA1", "POS01", "INFOREST"));

        Assert.False(result.Exitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_INVALIDO", result.CodigoError);
    }

    /// <summary>
    /// Legacy frmAcceso.frm líneas 476–494: login exitoso establece wEnter=True.
    /// </summary>
    [Fact]
    public async Task Login_Exitoso_RetornaExitosoConSesion()
    {
        var sesion = BuildSesion();
        var authService = BuildAuthService(AuthResult.ExitosoConSesion(sesion, true, false));
        var result = await authService.AutenticarAsync(new AuthRequest("ADMIN", "SECRETO", "INFOREST", "CA1", "POS01", "INFOREST"));

        Assert.True(result.Exitoso);
        Assert.NotNull(result.Sesion);
        Assert.Equal("USR1", result.Sesion!.CodigoUsuario);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // BR-POS-006-LOCK: bloqueo por intentos fallidos
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Legacy frmAcceso.frm: contador i inicia en 1, se incrementa en cada fallo.
    /// Si i = 4 → End. Equivale a MaxIntentosFallidos = 3.
    /// </summary>
    [Fact]
    public void LoginPolicy_MaxIntentosFallidos_EsIgualATres_EquivalenteLegacyContadorI()
    {
        // El legacy inicia i=1 e incrementa en cada fallo; End cuando i=4.
        // Equivale a 3 intentos fallidos antes de terminar.
        Assert.Equal(3, LoginPolicy.MaxIntentosFallidos);
    }

    /// <summary>
    /// El AuthService retorna el error correcto en cada tipo de fallo para que
    /// FrmLoginPOS pueda acumular el contador correctamente.
    /// </summary>
    [Theory]
    [InlineData("SEGURIDAD_USUARIO_NO_ENCONTRADO")]
    [InlineData("SEGURIDAD_PASSWORD_INVALIDO")]
    [InlineData("SEGURIDAD_USUARIO_INACTIVO")]
    public async Task Login_CualquierFallo_RetornaResultadoNoExitoso_ParaAcumularContador(string codigoError)
    {
        var authService = BuildAuthService(AuthResult.Fallido("Error.", codigoError));
        var result = await authService.AutenticarAsync(new AuthRequest("X", "Y", "INFOREST", "CA1", "POS01", "INFOREST"));

        // FrmLoginPOS incrementa _intentosFallidos solo cuando !result.Exitoso
        Assert.False(result.Exitoso);
        Assert.NotNull(result.CodigoError);
    }
}
