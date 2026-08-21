using System.Data;
using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Seguridad;
using Inforest.Infrastructure.Security;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace Inforest.Infrastructure.Tests.Security;

public sealed class AuthServiceTests
{
    [Fact]
    public async Task AutenticarAsync_UsuarioLegacyValido_CreaSesionYMigraPassword()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        var spExecutor = new Mock<ISpExecutor>();
        spExecutor
            .Setup(executor => executor.QueryAsync<LegacyModuleUserRecord>(
                connection.Object,
                "usp_Inforest_ObtieneUsuarios",
                It.IsAny<object>(),
                null,
                It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[]
            {
                new LegacyModuleUserRecord
                {
                    tCodigoUsuario = "USR01",
                    tGrupoUsuario = "01",
                    tDetallado = "Administrador",
                    tResumido = "ADMIN",
                    tPassword = LegacyPasswordCipher.Encrypt("SECRETO"),
                    lActivo = true,
                    tBandaMagnetica = string.Empty
                }
            });

        var rbac = new Mock<IRbacService>();
        rbac
            .Setup(service => service.ObtenerPermisosAsync("ADMIN", "02", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[]
            {
                new PermisoAcceso("ACC001", "01", "frmCaja", "btnCobrar", "Cobrar")
            });

        var sessionService = new SessionService();
        var audit = new Mock<IAuditoriaService>();
        audit
            .Setup(service => service.RegistrarIngresoAsync(It.IsAny<RegistroAccesoAuditoriaRequest>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(Result.Ok(99L));

        var hashStore = new Mock<IModernPasswordHashStore>();
        hashStore.Setup(store => store.GetHashAsync("USR01", It.IsAny<CancellationToken>())).ReturnsAsync((string?)null);
        hashStore.Setup(store => store.UpsertHashAsync("USR01", It.IsAny<string>(), It.IsAny<CancellationToken>())).ReturnsAsync(true);

        var service = new AuthService(
            connectionFactory.Object,
            spExecutor.Object,
            rbac.Object,
            sessionService,
            audit.Object,
            hashStore.Object,
            NullLogger<AuthService>.Instance);

        var result = await service.AutenticarAsync(new AuthRequest("admin", "SECRETO", "INFOREST", "01", "POS-01", "INFOREST"));

        Assert.True(result.Exitoso);
        Assert.NotNull(result.Sesion);
        Assert.True(result.PasswordMigrated);
        Assert.False(result.PasswordUpgradePending);
        Assert.Equal("USR01", result.Sesion!.CodigoUsuario);
        Assert.Equal(99L, result.Sesion.CorrelativoAcceso);
    }

    [Fact]
    public async Task AutenticarAsync_PasswordInvalido_Falla()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        var spExecutor = new Mock<ISpExecutor>();
        spExecutor
            .Setup(executor => executor.QueryAsync<LegacyModuleUserRecord>(
                connection.Object,
                "usp_Inforest_ObtieneUsuarios",
                It.IsAny<object>(),
                null,
                It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[]
            {
                new LegacyModuleUserRecord
                {
                    tCodigoUsuario = "USR01",
                    tGrupoUsuario = "01",
                    tDetallado = "Administrador",
                    tResumido = "ADMIN",
                    tPassword = LegacyPasswordCipher.Encrypt("SECRETO"),
                    lActivo = true,
                    tBandaMagnetica = string.Empty
                }
            });

        var service = new AuthService(
            connectionFactory.Object,
            spExecutor.Object,
            Mock.Of<IRbacService>(),
            new SessionService(),
            Mock.Of<IAuditoriaService>(),
            Mock.Of<IModernPasswordHashStore>(),
            NullLogger<AuthService>.Instance);

        var result = await service.AutenticarAsync(new AuthRequest("admin", "OTRA", "INFOREST", "01", "POS-01", "INFOREST"));

        Assert.False(result.Exitoso);
        Assert.Equal("SEGURIDAD_PASSWORD_INVALIDO", result.CodigoError);
    }
}
