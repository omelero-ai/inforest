using Inforest.Domain.Entities.Cocina;
using Inforest.Domain.Services;
using Inforest.Infrastructure.Kitchen;
using Moq;

namespace Inforest.Infrastructure.Tests.Kitchen;

/// <summary>
/// Pruebas de orquestación de cocina/comanda/KDS.
/// </summary>
public sealed class ProduccionCocinaServiceTests
{
    [Fact]
    public async Task EnviarNuevaOrdenAsync_ConKdsInforest_SincronizaPorArea()
    {
        var areaDispatcher = new Mock<IComandaAreaDispatcher>();
        var kdsDispatcher = new Mock<IKdsDispatcher>();
        var gateway = new Mock<IKdsLegacyGateway>();
        var service = new ProduccionCocinaService(new AreaProduccionRouter(), areaDispatcher.Object, kdsDispatcher.Object, gateway.Object);

        var item = ItemProduccion.Crear("PE000123", "1", "P01", "Pizza", 1, "001", true);
        var orden = OrdenProduccion.Crear("PE000123", "001", "M1", "Mesa 1", "Salon", "Mozo", "USR01", [item]);
        var config = ConfiguracionProduccionCocina.Crear(true, true, true, false, null, null, null, null);

        var result = await service.EnviarNuevaOrdenAsync(orden, config);

        Assert.True(result.EsExitoso);
        gateway.Verify(g => g.SincronizarOrdenInforestAsync("PE000123", 2, It.IsAny<CancellationToken>()), Times.Once);
        kdsDispatcher.Verify(
            d => d.EnviarNuevaOrdenAsync(
                It.IsAny<OrdenProduccion>(),
                It.IsAny<PlanProduccion>(),
                It.IsAny<ConfiguracionProduccionCocina>(),
                It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public async Task EnviarNuevaOrdenAsync_ConKdsXml_ImprimeYDespachaXml()
    {
        var areaDispatcher = new Mock<IComandaAreaDispatcher>();
        var kdsDispatcher = new Mock<IKdsDispatcher>();
        var gateway = new Mock<IKdsLegacyGateway>();
        var service = new ProduccionCocinaService(new AreaProduccionRouter(), areaDispatcher.Object, kdsDispatcher.Object, gateway.Object);

        var item = ItemProduccion.Crear("PE000123", "1", "P01", "Pizza", 1, "001", true, estacionKds: "001", canalKds: 0);
        var orden = OrdenProduccion.Crear("PE000123", "001", "M1", "Mesa 1", "Salon", "Mozo", "USR01", [item]);
        var config = ConfiguracionProduccionCocina.Crear(true, false, false, false, "/tmp/kds-a", null, null, null);

        var result = await service.EnviarNuevaOrdenAsync(orden, config);

        Assert.True(result.EsExitoso);
        areaDispatcher.Verify(d => d.ImprimirAsync(orden, It.IsAny<PlanProduccion>(), It.IsAny<CancellationToken>()), Times.Once);
        kdsDispatcher.Verify(
            d => d.EnviarNuevaOrdenAsync(
                orden,
                It.IsAny<PlanProduccion>(),
                config,
                It.IsAny<CancellationToken>()),
            Times.Once);
        gateway.Verify(g => g.SincronizarOrdenInforestAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<CancellationToken>()), Times.Never);
    }
}
