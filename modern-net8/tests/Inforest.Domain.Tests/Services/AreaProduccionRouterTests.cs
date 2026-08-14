using Inforest.Domain.Entities.Cocina;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Services;

namespace Inforest.Domain.Tests.Services;

/// <summary>
/// Pruebas de routing de cocina/KDS.
/// Legacy: frmVenta.frm + modKDS.bas + TPRODUCTOAREA.
/// </summary>
public sealed class AreaProduccionRouterTests
{
    private readonly AreaProduccionRouter _router = new();

    [Fact]
    public void CrearPlan_AgrupaComandasPorAreaYCanalKds()
    {
        var orden = OrdenProduccion.Crear(
            "PE000123",
            "001",
            "M01",
            "Mesa 1",
            "Salon",
            "MOZO",
            "USR01",
            [
                ItemProduccion.Crear("PE000123", "1", "P01", "Pizza", 1, "001", true, estacionKds: "001", canalKds: 0),
                ItemProduccion.Crear("PE000123", "2", "P02", "Pasta", 1, "002", true, estacionKds: "002", canalKds: 1)
            ]);

        var config = ConfiguracionProduccionCocina.Crear(
            kdsHabilitado: true,
            kdsInforestHabilitado: false,
            kdsPorAreaHabilitado: true,
            kdsTodosHabilitado: false,
            directorioOrderInfoPrimario: "/tmp/kds-a",
            directorioOrderInfoSecundario: "/tmp/kds-b",
            directorioBumpPrimario: null,
            directorioBumpSecundario: null);

        var plan = _router.CrearPlan(orden, config);

        Assert.Equal(2, plan.ComandasPorArea.Count);
        Assert.Equal(2, plan.CanalesKds.Count);
        Assert.Contains(plan.ComandasPorArea, area => area.CodigoArea == "001");
        Assert.Contains(plan.ComandasPorArea, area => area.CodigoArea == "002");
        Assert.Contains(plan.CanalesKds, canal => canal.Canal == 0);
        Assert.Contains(plan.CanalesKds, canal => canal.Canal == 1);
    }

    [Fact]
    public void CrearPlan_ItemImprimibleSinArea_LanzaExcepcion()
    {
        var orden = OrdenProduccion.Crear(
            "PE000123",
            "001",
            null,
            null,
            "Salon",
            "MOZO",
            "USR01",
            [ItemProduccion.Crear("PE000123", "1", "P01", "Pizza", 1, null, true)]);

        var config = ConfiguracionProduccionCocina.Crear(false, false, false, false, null, null, null, null);

        var ex = Assert.Throws<DomainException>(() => _router.CrearPlan(orden, config));

        Assert.Equal("COCINA_AREA_REQUERIDA", ex.Code);
    }
}
