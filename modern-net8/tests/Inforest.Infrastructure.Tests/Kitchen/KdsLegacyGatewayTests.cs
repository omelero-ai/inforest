using System.Data;
using Inforest.Application.Interfaces;
using Inforest.Infrastructure.Kitchen;
using Inforest.Infrastructure.Kitchen.Models;
using Moq;

namespace Inforest.Infrastructure.Tests.Kitchen;

/// <summary>
/// Pruebas de gateway Legacy KDS.
/// Verifican nombres de SP y parámetros críticos.
/// </summary>
public sealed class KdsLegacyGatewayTests
{
    [Fact]
    public async Task ObtenerConfiguracionAsync_UsaStoredProcedureLegacy()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        var executor = new Mock<ISpExecutor>();

        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        executor
            .Setup(sp => sp.QueryFirstOrDefaultAsync<KdsPathConfigurationRecord>(
                connection.Object,
                "USP_KDS_ObtenerPath",
                null,
                null,
                It.IsAny<CancellationToken>()))
            .ReturnsAsync(new KdsPathConfigurationRecord("/tmp/a", null, null, null, null, null, false, false, false, true));

        var gateway = new KdsLegacyGateway(connectionFactory.Object, executor.Object);

        var result = await gateway.ObtenerConfiguracionAsync();

        Assert.NotNull(result);
        executor.VerifyAll();
    }

    [Fact]
    public async Task ObtenerDetallePedidoAsync_ConKdsX_UsaStoredProcedureExtendido()
    {
        var connectionFactory = new Mock<IDbConnectionFactory>();
        var connection = new Mock<IDbConnection>();
        var executor = new Mock<ISpExecutor>();

        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(connection.Object);

        executor
            .Setup(sp => sp.QueryAsync<KdsDetallePedidoRecord>(
                connection.Object,
                "USP_KDS_ObtenerDetallePedidox",
                It.Is<object>(param => ParamValue<int>(param, "accion") == 2 && ParamValue<int>(param, "xkds") == 1),
                null,
                It.IsAny<CancellationToken>()))
            .ReturnsAsync([]);

        var gateway = new KdsLegacyGateway(connectionFactory.Object, executor.Object);

        await gateway.ObtenerDetallePedidoAsync("PE000123", 2, 1);

        executor.VerifyAll();
    }

    [Fact]
    public async Task SincronizarOrdenInforestAsync_UsaStoredProcedureDeActualizacion()
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
                "KDS_Usp_spUpdate_DPEDIDO_Orden",
                It.Is<object>(param => ParamValue<int>(param, "tipo") == 2 && ParamValue<string>(param, "pedido") == "PE000123"),
                null,
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var gateway = new KdsLegacyGateway(connectionFactory.Object, executor.Object);

        await gateway.SincronizarOrdenInforestAsync("PE000123", 2);

        executor.VerifyAll();
    }

    private static T ParamValue<T>(object param, string property)
    {
        var value = param.GetType().GetProperty(property)?.GetValue(param);
        return (T)value!;
    }
}
