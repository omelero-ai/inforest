using Inforest.Application.Caja;
using Inforest.Domain.Entities.Caja;
using Moq;

namespace Inforest.Application.Tests.Caja;

public class ObtenerMediosPagoHandlerTests
{
    [Fact]
    public async Task HandleAsync_DevuelveMediosPagoActivos()
    {
        var medios = new List<MedioPago>
        {
            MedioPago.Crear("01", "Efectivo", TipoMedioPago.Efectivo),
            MedioPago.Crear("02", "Tarjeta", TipoMedioPago.Tarjeta)
        };

        var repo = new Mock<IMedioPagoRepository>();
        repo.Setup(r => r.ObtenerActivosAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(medios.AsReadOnly());

        var handler = new ObtenerMediosPagoHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerMediosPagoQuery());

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count);
        Assert.Contains(result.Valor, m => m.Codigo == "01");
        Assert.Contains(result.Valor, m => m.Tipo == TipoMedioPago.Tarjeta);
    }

    [Fact]
    public async Task HandleAsync_SinMedios_DevuelveListaVacia()
    {
        var repo = new Mock<IMedioPagoRepository>();
        repo.Setup(r => r.ObtenerActivosAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(Array.Empty<MedioPago>());

        var handler = new ObtenerMediosPagoHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerMediosPagoQuery());

        Assert.True(result.EsExitoso);
        Assert.Empty(result.Valor!);
    }
}
