using Inforest.Application.Interfaces;
using Inforest.Application.Seguridad;
using Inforest.Domain.Common;
using Moq;

namespace Inforest.Application.Tests.Seguridad;

/// <summary>
/// Tests para CambiarPasswordHandler.
/// Legacy: frmCambiarContrasenia.frm CmdAceptar_Click. BR-006.
/// </summary>
public class CambiarPasswordHandlerTests
{
    private static CambiarPasswordHandler BuildHandler(Mock<IAuthService>? authMock = null)
    {
        authMock ??= new Mock<IAuthService>();
        return new CambiarPasswordHandler(authMock.Object);
    }

    // ── Validaciones de entrada ────────────────────────────────────────────────

    [Fact]
    public async Task CambiarPassword_UsuarioVacio_Falla()
    {
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new CambiarPasswordCommand("", "old", "new1", "new1"));

        Assert.False(result.EsExitoso);
        Assert.Equal("SEGURIDAD_USUARIO_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task CambiarPassword_PasswordActualVacio_Falla()
    {
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new CambiarPasswordCommand("user", "", "new1", "new1"));

        Assert.False(result.EsExitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_ACTUAL_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task CambiarPassword_PasswordNuevoVacio_Falla()
    {
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new CambiarPasswordCommand("user", "old", "", ""));

        Assert.False(result.EsExitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_NUEVO_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task CambiarPassword_ConfirmacionVacia_Falla()
    {
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new CambiarPasswordCommand("user", "old", "new1", ""));

        Assert.False(result.EsExitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_CONFIRMACION_REQUERIDA", result.CodigoError);
    }

    [Theory]
    [InlineData("old", "old")]   // mismo valor
    [InlineData("OLD", "old")]   // mismo valor, diferente case
    public async Task CambiarPassword_NuevoIgualAlActual_Falla(string actual, string nuevo)
    {
        // Legacy: "La nueva contraseña debe ser distinta a la actual"
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new CambiarPasswordCommand("user", actual, nuevo, nuevo));

        Assert.False(result.EsExitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_IGUAL_A_ACTUAL", result.CodigoError);
    }

    [Fact]
    public async Task CambiarPassword_ConfirmacionNoCoincide_Falla()
    {
        // Legacy: "Nueva contraseña no coincide con la confirmación"
        var handler = BuildHandler();
        var result = await handler.HandleAsync(new CambiarPasswordCommand("user", "old", "new1", "new2"));

        Assert.False(result.EsExitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_CONFIRMACION_NO_COINCIDE", result.CodigoError);
    }

    [Fact]
    public async Task CambiarPassword_Exitoso_LlamaAlServicioYRetornaOk()
    {
        var authMock = new Mock<IAuthService>();
        authMock
            .Setup(a => a.CambiarPasswordAsync("user", "oldPwd", "newPwd", It.IsAny<CancellationToken>()))
            .ReturnsAsync(Result.Ok());

        var handler = BuildHandler(authMock);
        var result = await handler.HandleAsync(new CambiarPasswordCommand("user", "oldPwd", "newPwd", "newPwd"));

        Assert.True(result.EsExitoso);
        authMock.Verify(a => a.CambiarPasswordAsync("user", "oldPwd", "newPwd", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task CambiarPassword_ServicioFalla_PropagaError()
    {
        var authMock = new Mock<IAuthService>();
        authMock
            .Setup(a => a.CambiarPasswordAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(Result.Fail("Contraseña actual incorrecta.", "SEGURIDAD_PASSWORD_ACTUAL_INVALIDO"));

        var handler = BuildHandler(authMock);
        var result = await handler.HandleAsync(new CambiarPasswordCommand("user", "wrong", "newPwd", "newPwd"));

        Assert.False(result.EsExitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_ACTUAL_INVALIDO", result.CodigoError);
    }
}
