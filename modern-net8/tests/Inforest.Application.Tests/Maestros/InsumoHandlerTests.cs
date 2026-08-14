using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;
using Moq;

namespace Inforest.Application.Tests.Maestros;

/// <summary>
/// Tests de handlers POS-FUNC-013 — Insumos/descargo.
/// BR-INSUMO-001, BR-INSUMO-002, BR-INSUMO-003, BR-INSUMO-004.
/// </summary>
public class InsumoHandlerTests
{
    // ── ListarInsumosHandler ──────────────────────────────────────────────────

    [Fact]
    public async Task ListarInsumos_RetornaListaCompleta()
    {
        var i1 = Insumo.Crear("00000001", "SAL", "user", 10, true);
        var i2 = Insumo.Crear("00000002", "ACEITE", "user", 5, true);
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ObtenerTodosAsync(true, default))
            .ReturnsAsync(new List<Insumo> { i1, i2 });

        var handler = new ListarInsumosHandler(repo.Object);
        var result  = await handler.HandleAsync(new ListarInsumosQuery(true));

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count);
    }

    [Fact]
    public async Task ListarInsumos_ListaVacia_RetornaExitoConCeroItems()
    {
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ObtenerTodosAsync(true, default))
            .ReturnsAsync(new List<Insumo>());

        var handler = new ListarInsumosHandler(repo.Object);
        var result  = await handler.HandleAsync(new ListarInsumosQuery(true));

        Assert.True(result.EsExitoso);
        Assert.Empty(result.Valor!);
    }

    // ── ObtenerInsumoPorCodigoHandler ─────────────────────────────────────────

    [Fact]
    public async Task ObtenerPorCodigo_Existente_RetornaInsumo()
    {
        var insumo = Insumo.Crear("00000001", "SAL", "user", 10, true);
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("00000001", default))
            .ReturnsAsync(insumo);

        var handler = new ObtenerInsumoPorCodigoHandler(repo.Object);
        var result  = await handler.HandleAsync(new ObtenerInsumoPorCodigoQuery("00000001"));

        Assert.True(result.EsExitoso);
        Assert.Equal("SAL", result.Valor!.Descripcion);
    }

    [Fact]
    public async Task ObtenerPorCodigo_Inexistente_RetornaFallo()
    {
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("NOEXISTE", default))
            .ReturnsAsync((Insumo?)null);

        var handler = new ObtenerInsumoPorCodigoHandler(repo.Object);
        var result  = await handler.HandleAsync(new ObtenerInsumoPorCodigoQuery("NOEXISTE"));

        Assert.False(result.EsExitoso);
        Assert.Equal("INSUMO_NO_ENCONTRADO", result.CodigoError);
    }

    // ── AgregarInsumoHandler ──────────────────────────────────────────────────

    [Fact]
    public async Task AgregarInsumo_DescripcionNueva_Agrega()
    {
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ExisteDescripcionAsync("SAL", default)).ReturnsAsync(false);
        repo.Setup(r => r.ObtenerProximoCodigoAsync(default)).ReturnsAsync("00000001");
        repo.Setup(r => r.AgregarAsync(It.IsAny<Insumo>(), "C01", true, default))
            .Returns(Task.CompletedTask);

        var handler = new AgregarInsumoHandler(repo.Object);
        var result  = await handler.HandleAsync(new AgregarInsumoCommand("sal", "user", 5.0, true, "C01"));

        Assert.True(result.EsExitoso);
        Assert.Equal("00000001", result.Valor);
        repo.Verify(r => r.AgregarAsync(It.Is<Insumo>(i => i.Descripcion == "SAL"), "C01", true, default), Times.Once);
    }

    [Fact]
    public async Task AgregarInsumo_DescripcionDuplicada_RetornaFallo()
    {
        // BR-INSUMO-001: no duplicados
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ExisteDescripcionAsync("SAL", default)).ReturnsAsync(true);

        var handler = new AgregarInsumoHandler(repo.Object);
        var result  = await handler.HandleAsync(new AgregarInsumoCommand("sal", "user", 0, true, "C01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("INSUMO_DESCRIPCION_DUPLICADA", result.CodigoError);
        repo.Verify(r => r.AgregarAsync(It.IsAny<Insumo>(), It.IsAny<string>(), It.IsAny<bool>(), default), Times.Never);
    }

    [Fact]
    public async Task AgregarInsumo_DescripcionEnMayusculas_BR_INSUMO_002()
    {
        // BR-INSUMO-002: la descripción debe almacenarse en mayúsculas
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ExisteDescripcionAsync("aceite oliva", default)).ReturnsAsync(false);
        repo.Setup(r => r.ObtenerProximoCodigoAsync(default)).ReturnsAsync("00000002");
        Insumo? capturado = null;
        repo.Setup(r => r.AgregarAsync(It.IsAny<Insumo>(), It.IsAny<string>(), true, default))
            .Callback<Insumo, string, bool, CancellationToken>((i, _, _, _) => capturado = i)
            .Returns(Task.CompletedTask);

        var handler = new AgregarInsumoHandler(repo.Object);
        await handler.HandleAsync(new AgregarInsumoCommand("aceite oliva", "user", 0, true, "C01"));

        Assert.NotNull(capturado);
        Assert.Equal("ACEITE OLIVA", capturado!.Descripcion);
    }

    // ── ModificarInsumoHandler ────────────────────────────────────────────────

    [Fact]
    public async Task ModificarInsumo_Existente_Modifica()
    {
        var insumo = Insumo.Crear("00000001", "SAL", "user", 10, true);
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("00000001", default)).ReturnsAsync(insumo);
        repo.Setup(r => r.ModificarAsync(It.IsAny<Insumo>(), "C01", true, default)).Returns(Task.CompletedTask);

        var handler = new ModificarInsumoHandler(repo.Object);
        var result  = await handler.HandleAsync(
            new ModificarInsumoCommand("00000001", "sal gruesa", "user", 20, true, true, "C01"));

        Assert.True(result.EsExitoso);
        Assert.Equal("SAL GRUESA", insumo.Descripcion);
    }

    [Fact]
    public async Task ModificarInsumo_NoExistente_RetornaFallo()
    {
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("NOEXISTE", default)).ReturnsAsync((Insumo?)null);

        var handler = new ModificarInsumoHandler(repo.Object);
        var result  = await handler.HandleAsync(
            new ModificarInsumoCommand("NOEXISTE", "desc", "user", 0, true, true, "C01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("INSUMO_NO_ENCONTRADO", result.CodigoError);
    }

    // ── EliminarInsumoHandler ─────────────────────────────────────────────────

    [Fact]
    public async Task EliminarInsumo_Existente_Elimina()
    {
        var insumo = Insumo.Crear("00000001", "SAL", "user");
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("00000001", default)).ReturnsAsync(insumo);
        repo.Setup(r => r.EliminarAsync("00000001", default)).Returns(Task.CompletedTask);

        var handler = new EliminarInsumoHandler(repo.Object);
        var result  = await handler.HandleAsync(new EliminarInsumoCommand("00000001"));

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.EliminarAsync("00000001", default), Times.Once);
    }

    [Fact]
    public async Task EliminarInsumo_NoExistente_RetornaFallo()
    {
        // BR-INSUMO-004: no se elimina si no existe
        var repo = new Mock<IInsumoRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("NOEXISTE", default)).ReturnsAsync((Insumo?)null);

        var handler = new EliminarInsumoHandler(repo.Object);
        var result  = await handler.HandleAsync(new EliminarInsumoCommand("NOEXISTE"));

        Assert.False(result.EsExitoso);
        Assert.Equal("INSUMO_NO_ENCONTRADO", result.CodigoError);
        repo.Verify(r => r.EliminarAsync(It.IsAny<string>(), default), Times.Never);
    }
}
