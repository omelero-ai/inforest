using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Domain.Tests.Maestros;

/// <summary>
/// Tests del maestro de productos.
/// Legacy: TPRODUCTO / BR-001.
/// </summary>
public class ProductoMaestroTests
{
    [Fact]
    public void ObtenerPrecioPorCanal_UsaPrecioEspecificoSiExiste()
    {
        var producto = ProductoMaestro.Crear(
            "PROD01", "BEB", "Gaseosa", "tester", 10m,
            precioDelivery: 12m,
            precioLlevar: 11m);

        Assert.Equal(10m, producto.ObtenerPrecioPorCanal(CanalVenta.Local));
        Assert.Equal(12m, producto.ObtenerPrecioPorCanal(CanalVenta.Delivery));
        Assert.Equal(11m, producto.ObtenerPrecioPorCanal(CanalVenta.Llevar));
    }

    [Fact]
    public void ObtenerPrecioPorCanal_SiCanalSinPrecio_UsaBase()
    {
        var producto = ProductoMaestro.Crear("PROD01", "BEB", "Gaseosa", "tester", 10m);

        Assert.Equal(10m, producto.ObtenerPrecioPorCanal(CanalVenta.Delivery));
        Assert.Equal(10m, producto.ObtenerPrecioPorCanal(CanalVenta.Llevar));
    }
}
