using System.Xml.Linq;
using Inforest.Domain.Entities.Cocina;
using Inforest.Infrastructure.Kitchen;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace Inforest.Infrastructure.Tests.Kitchen;

/// <summary>
/// Pruebas de generación y procesamiento XML KDS.
/// Legacy: modKDS.bas.
/// </summary>
public sealed class KdsXmlDispatcherTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"inforest-kds-{Guid.NewGuid():N}");

    [Fact]
    public async Task EnviarNuevaOrdenAsync_GeneraXmlConComboYCondimentos()
    {
        Directory.CreateDirectory(_root);
        var gateway = new Mock<IKdsLegacyGateway>();
        var dispatcher = new KdsXmlDispatcher(gateway.Object, NullLogger<KdsXmlDispatcher>.Instance);

        var item = ItemProduccion.Crear(
            "PE000123",
            "1",
            "P01",
            "Hamburguesa",
            2,
            "007",
            true,
            itemCombo: "01",
            estacionKds: "007",
            canalKds: 0,
            esCombo: true,
            nombreCombo: "Combo Lunch",
            condimentos: [new CondimentoProduccion("10", "+ Queso")]);

        var orden = OrdenProduccion.Crear("PE000123", "001", "M01", "Mesa 1", "Salon", "Mozo 1", "USR01", [item]);
        var plan = new PlanProduccion([], [new KdsCanalPlan(0, [item])], false);
        var config = ConfiguracionProduccionCocina.Crear(true, false, false, false, _root, null, null, null);

        await dispatcher.EnviarNuevaOrdenAsync(orden, plan, config);

        var path = Path.Combine(_root, "123.xml");
        Assert.True(File.Exists(path));

        var xml = XDocument.Load(path);
        Assert.Equal("CLHamburguesa", xml.Root?.Descendants("Name").FirstOrDefault()?.Value);
        Assert.Equal("007", xml.Root?.Descendants("KDSStation").FirstOrDefault()?.Value);
        Assert.Equal("+ Queso", xml.Root?.Descendants("Condiment").Elements("Name").FirstOrDefault()?.Value);
    }

    [Fact]
    public async Task ProcesarBumpNotificationsAsync_RegistraSalidaYMueveArchivo()
    {
        var bump = Path.Combine(_root, "bump");
        Directory.CreateDirectory(bump);
        var gateway = new Mock<IKdsLegacyGateway>();
        var dispatcher = new KdsXmlDispatcher(gateway.Object, NullLogger<KdsXmlDispatcher>.Instance);

        await File.WriteAllTextAsync(
            Path.Combine(bump, "order.xml"),
            """
            <Transaction>
              <Order>
                <ID>123</ID>
                <Item>
                  <ID>1</ID>
                </Item>
              </Order>
            </Transaction>
            """);

        var config = ConfiguracionProduccionCocina.Crear(true, false, false, false, _root, null, bump, null);

        var processed = await dispatcher.ProcesarBumpNotificationsAsync(config);

        Assert.Equal(1, processed);
        gateway.Verify(g => g.RegistrarTiempoSalidaAsync("123", "1", It.IsAny<DateTime>(), It.IsAny<CancellationToken>()), Times.Once);
        Assert.True(File.Exists(Path.Combine(bump, "Historial", "order.xml")));
    }

    public void Dispose()
    {
        if (Directory.Exists(_root))
            Directory.Delete(_root, recursive: true);
    }
}
