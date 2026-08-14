using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;
using Moq;

namespace Inforest.Application.Tests.Ventas;

/// <summary>
/// Tests para CorrelativoHandlers.
/// Legacy: modPuntoVenta.bas Sub Main — NFactura = tUltimoNumero o "Sin Correlativo". BR-POS-008.
/// </summary>
public class CorrelativoHandlersTests
{
    private static CorrelativoDocumento BuildCorrelativo(
        string caja = "001",
        string tipo = "01",
        string? ultimoNumero = "000000120")
        => CorrelativoDocumento.Crear(
            caja, "IMP01", tipo,
            "Factura", "F001", "F001",
            ultimoNumero, null,
            null, null,
            false, true, false, false, false, false, false, false,
            "USR01", DateTime.Today);

    // ─────────────────────────────────────────────────────────────────────────
    // ObtenerCorrelativosPorCajaHandler
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task ObtenerPorCaja_CajaValida_RetornaLista()
    {
        var repo = new Mock<ICorrelativoRepository>();
        repo.Setup(r => r.ObtenerPorCajaAsync("001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { BuildCorrelativo() });

        var handler = new ObtenerCorrelativosPorCajaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerCorrelativosPorCajaQuery("001"));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!);
    }

    [Fact]
    public async Task ObtenerPorCaja_CajaVacia_Falla()
    {
        var repo = new Mock<ICorrelativoRepository>();
        var handler = new ObtenerCorrelativosPorCajaHandler(repo.Object);

        var result = await handler.HandleAsync(new ObtenerCorrelativosPorCajaQuery(""));

        Assert.False(result.EsExitoso);
        Assert.Equal("CORRELATIVO_CAJA_REQUERIDA", result.CodigoError);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // ObtenerCorrelativoFacturaHandler — BR-POS-008
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task ObtenerCorrelativoFactura_Existe_RetornaNumero()
    {
        var repo = new Mock<ICorrelativoRepository>();
        repo.Setup(r => r.ObtenerPorCajaYTipoAsync("001", "01", It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildCorrelativo(ultimoNumero: "000000120"));

        var handler = new ObtenerCorrelativoFacturaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerCorrelativoFacturaQuery("001", "01"));

        Assert.True(result.EsExitoso);
        Assert.Equal("000000120", result.Valor);
    }

    [Fact]
    public async Task ObtenerCorrelativoFactura_NoExiste_RetornaSinCorrelativo()
    {
        // BR-POS-008: si no hay registro en TTIPODOCUMENTOIMPRESORA → "Sin Correlativo"
        var repo = new Mock<ICorrelativoRepository>();
        repo.Setup(r => r.ObtenerPorCajaYTipoAsync("001", "01", It.IsAny<CancellationToken>()))
            .ReturnsAsync((CorrelativoDocumento?)null);

        var handler = new ObtenerCorrelativoFacturaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerCorrelativoFacturaQuery("001", "01"));

        Assert.True(result.EsExitoso);
        Assert.Equal("Sin Correlativo", result.Valor);
    }

    [Fact]
    public async Task ObtenerCorrelativoFactura_NumeroNulo_RetornaSinCorrelativo()
    {
        var repo = new Mock<ICorrelativoRepository>();
        repo.Setup(r => r.ObtenerPorCajaYTipoAsync("001", "01", It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildCorrelativo(ultimoNumero: null));

        var handler = new ObtenerCorrelativoFacturaHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerCorrelativoFacturaQuery("001", "01"));

        Assert.True(result.EsExitoso);
        Assert.Equal("Sin Correlativo", result.Valor);
    }

    [Fact]
    public async Task ObtenerCorrelativoFactura_TipoEmisionVacio_Falla()
    {
        var repo = new Mock<ICorrelativoRepository>();
        var handler = new ObtenerCorrelativoFacturaHandler(repo.Object);

        var result = await handler.HandleAsync(new ObtenerCorrelativoFacturaQuery("001", ""));

        Assert.False(result.EsExitoso);
        Assert.Equal("CORRELATIVO_TIPO_EMISION_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerTodosCorrelativos_RetornaLista()
    {
        var repo = new Mock<ICorrelativoRepository>();
        repo.Setup(r => r.ObtenerTodosAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { BuildCorrelativo("001", "01"), BuildCorrelativo("002", "03") });

        var handler = new ObtenerTodosCorrelativosHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerTodosCorrelativosQuery());

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Domain entity — CorrelativoDocumento
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public void CorrelativoDocumento_TieneNumero_EsVerdadero()
    {
        var c = BuildCorrelativo(ultimoNumero: "000000100");
        Assert.True(c.TieneCorrelativo);
        Assert.Equal("000000100", c.ObtenerNumeroActual());
    }

    [Fact]
    public void CorrelativoDocumento_SinCorrelativoLiteral_NotieneCorrelativo()
    {
        var c = BuildCorrelativo(ultimoNumero: "Sin Correlativo");
        Assert.False(c.TieneCorrelativo);
        Assert.Equal("Sin Correlativo", c.ObtenerNumeroActual());
    }

    [Fact]
    public void CorrelativoDocumento_NumeroNulo_NoTieneCorrelativo()
    {
        var c = BuildCorrelativo(ultimoNumero: null);
        Assert.False(c.TieneCorrelativo);
        Assert.Equal("Sin Correlativo", c.ObtenerNumeroActual());
    }

    [Fact]
    public void CorrelativoDocumento_SinFechaCaducidad_EsVigente()
    {
        var c = BuildCorrelativo();
        Assert.True(c.EsVigente);
    }

    [Fact]
    public void CorrelativoDocumento_FechaCaducidadVencida_NoEsVigente()
    {
        var c = CorrelativoDocumento.Crear(
            "001", "IMP01", "01",
            "Factura", "F001", "F001",
            "000000100", null,
            null, DateTime.Today.AddDays(-1),
            false, true, false, false, false, false, false, false,
            "USR01", DateTime.Today);

        Assert.False(c.EsVigente);
    }
}
