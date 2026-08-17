using Inforest.Application.Maestros;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Maestros;
using Moq;

namespace Inforest.Application.Tests.Ventas;

public class SolicitudBoletaHandlersTests
{
    [Fact]
    public async Task ObtenerClienteGeneral_SiExiste_RetornaExistente()
    {
        var cliente = Cliente.Crear("00015", "Cliente General", "00000000", "Lima", "ADMIN", tipoIdentidad: "01");
        var repo = new Mock<IClienteRepository>();
        repo.Setup(r => r.ObtenerPorIdentidadAsync("00000000", It.IsAny<CancellationToken>()))
            .ReturnsAsync(cliente);

        var handler = new ObtenerClienteGeneralBoletaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerClienteGeneralBoletaQuery());

        Assert.True(result.EsExitoso);
        Assert.Equal("00015", result.Valor!.CodigoCliente);
        repo.Verify(r => r.InsertarAsync(It.IsAny<Cliente>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ObtenerClienteGeneral_SiNoExiste_CreaRegistroBase()
    {
        var repo = new Mock<IClienteRepository>();
        repo.Setup(r => r.ObtenerPorIdentidadAsync("00000000", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Cliente?)null);
        repo.Setup(r => r.ObtenerProximoCodigoAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync("00021");
        repo.Setup(r => r.InsertarAsync(It.IsAny<Cliente>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = new ObtenerClienteGeneralBoletaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerClienteGeneralBoletaQuery());

        Assert.True(result.EsExitoso);
        Assert.Equal("00021", result.Valor!.CodigoCliente);
        Assert.Equal("00000000", result.Valor.Identidad);
        repo.Verify(r => r.InsertarAsync(It.Is<Cliente>(c =>
            c.CodigoCliente == "00021" &&
            c.Empresa == "Cliente General" &&
            c.Identidad == "00000000"), It.IsAny<CancellationToken>()), Times.Once);
    }
}
