using System.Data;
using Inforest.Application.Interfaces;
using Inforest.Infrastructure.Kitchen;
using Moq;

namespace Inforest.Infrastructure.Tests.Kitchen;

public sealed class MensajeCocinaRepositoryTests
{
    [Fact]
    public async Task ObtenerMensajes_UsaStoredProcedureLegacy()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        var executor = new Mock<ISpExecutor>();

        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        executor
            .Setup(sp => sp.QueryAsync<MensajeCocinaRepository.MensajeCocinaListRow>(
                connection.Object,
                "USP_LISTARMENSAJES",
                It.IsAny<object?>(),
                null,
                It.IsAny<CancellationToken>()))
            .ReturnsAsync([]);

        var repository = new MensajeCocinaRepository(connectionFactory.Object, executor.Object);
        await repository.ObtenerMensajesAsync(DateTime.Today, DateTime.Today, string.Empty);

        executor.Verify(sp => sp.QueryAsync<MensajeCocinaRepository.MensajeCocinaListRow>(
            connection.Object,
            "USP_LISTARMENSAJES",
            It.IsAny<object?>(),
            null,
            It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Agregar_UsaStoredProcedureLegacy()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        var executor = new Mock<ISpExecutor>();

        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        executor
            .Setup(sp => sp.ExecuteAsync(
                connection.Object,
                "USP_AGREGARMENSAJE",
                It.Is<object>(param => ParamValue<string>(param, "codigo") == "26000001" && ParamValue<string>(param, "tcaja") == "C01"),
                null,
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var repository = new MensajeCocinaRepository(connectionFactory.Object, executor.Object);
        var mensaje = Inforest.Domain.Entities.Cocina.MensajeCocina.Crear("26000001", "USR1", "mensaje", "C01", true, true);
        await repository.AgregarAsync(mensaje);

        executor.VerifyAll();
    }

    [Fact]
    public async Task Modificar_UsaStoredProcedureLegacy()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        var executor = new Mock<ISpExecutor>();

        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        executor
            .Setup(sp => sp.ExecuteAsync(
                connection.Object,
                "USP_MODIFICARMENSAJE",
                It.Is<object>(param => ParamValue<string>(param, "codigo") == "26000001" && ParamValue<bool>(param, "lactivaralerta")),
                null,
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var repository = new MensajeCocinaRepository(connectionFactory.Object, executor.Object);
        var mensaje = Inforest.Domain.Entities.Cocina.MensajeCocina.Crear("26000001", "USR1", "mensaje", "C01", true, false);
        mensaje.Actualizar("USR2", "nuevo", "C01", true, true);
        await repository.ModificarAsync(mensaje);

        executor.VerifyAll();
    }

    [Fact]
    public async Task Eliminar_UsaStoredProcedureLegacy()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        var executor = new Mock<ISpExecutor>();

        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        executor
            .Setup(sp => sp.ExecuteAsync(
                connection.Object,
                "USP_ELIMINARRMENSAJES",
                It.Is<object>(param => ParamValue<string>(param, "tCODIGO") == "26000001"),
                null,
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var repository = new MensajeCocinaRepository(connectionFactory.Object, executor.Object);
        await repository.EliminarAsync("26000001");

        executor.VerifyAll();
    }

    [Fact]
    public async Task CerrarActivosPorCaja_UsaStoredProcedureLegacy()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        var executor = new Mock<ISpExecutor>();

        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        executor
            .Setup(sp => sp.ExecuteAsync(
                connection.Object,
                "USP_CERRAR_MENSAJES_CIERRETURNO",
                It.Is<object>(param => ParamValue<string>(param, "usuario") == "USR1" && ParamValue<string>(param, "tcaja") == "C01"),
                null,
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var repository = new MensajeCocinaRepository(connectionFactory.Object, executor.Object);
        await repository.CerrarActivosPorCajaAsync("USR1", "C01");

        executor.VerifyAll();
    }

    private static T ParamValue<T>(object param, string property)
    {
        var value = param.GetType().GetProperty(property)?.GetValue(param);
        return (T)value!;
    }
}
