using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;
using Moq;

namespace Inforest.Infrastructure.Tests.Maestros;

/// <summary>
/// Tests de contrato del repositorio de grupos de producto usando Moq.
/// </summary>
public class GrupoProductoRepositoryTests
{
    [Fact]
    public async Task ObtenerTodos_RetornaListaConfigurada()
    {
        var esperado = new List<GrupoProducto>
        {
            GrupoProducto.Crear("BEB", "Bebidas", "BEB", "tester"),
            GrupoProducto.Crear("COM", "Comidas", "COM", "tester")
        };

        var mock = new Mock<IGrupoProductoRepository>();
        mock.Setup(r => r.ObtenerTodosAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(esperado);

        var resultado = await mock.Object.ObtenerTodosAsync();

        Assert.Equal(2, resultado.Count);
        Assert.Equal("BEB", resultado[0].CodigoGrupo);
    }

    [Fact]
    public async Task ObtenerPorCodigo_InvocaRepositorioConCodigoEsperado()
    {
        var grupo = GrupoProducto.Crear("BEB", "Bebidas", "BEB", "tester");
        var mock = new Mock<IGrupoProductoRepository>();
        mock.Setup(r => r.ObtenerPorCodigoAsync("BEB", It.IsAny<CancellationToken>()))
            .ReturnsAsync(grupo);

        var resultado = await mock.Object.ObtenerPorCodigoAsync("BEB");

        Assert.NotNull(resultado);
        Assert.Equal("BEB", resultado!.CodigoGrupo);
        mock.Verify(r => r.ObtenerPorCodigoAsync("BEB", It.IsAny<CancellationToken>()), Times.Once);
    }
}
