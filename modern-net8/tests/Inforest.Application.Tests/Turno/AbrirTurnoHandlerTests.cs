using Inforest.Application.Configuracion;
using Inforest.Application.Turno;
using Moq;

namespace Inforest.Application.Tests.Turno;

/// <summary>
/// Tests para AbrirTurnoHandler.
/// Legacy: frmInicio.frm — cmdOpcion_Click(0).
/// Reglas: BR-TURNO-001 (apertura nueva), BR-TURNO-002 (re-apertura), BR-TC-001.
/// </summary>
public class AbrirTurnoHandlerTests
{
    // ─────────────────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────────────────

    private static IDiaContableService DiaContableService()
    {
        var mock = new Mock<IDiaContableService>();
        mock.Setup(s => s.ObtenerDiaContableAsync(
                It.IsAny<bool>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DateTime(2026, 8, 21));
        return mock.Object;
    }

    private static TurnoExistente BuildTurnoExistente(string codigo = "2600000001", bool cerrado = false)
        => new(codigo, "CA1", "USR1", new DateTime(2026, 8, 21), cerrado, 100m, 0m);

    private static ITurnoRepository TurnoRepoSinTurnoExistente(string correlativo = "2600000001")
    {
        var mock = new Mock<ITurnoRepository>();
        mock.Setup(r => r.ObtenerUltimoTurnoAsync(
                It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<ModoConsultaTurno>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((TurnoExistente?)null);
        mock.Setup(r => r.GenerarCorrelativoAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(correlativo);
        mock.Setup(r => r.InsertarAsync(It.IsAny<Domain.Entities.Configuracion.Turno>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);
        return mock.Object;
    }

    private static ITurnoRepository TurnoRepoConTurnoAbierto(string codigoTurno = "2600000001")
    {
        var mock = new Mock<ITurnoRepository>();
        mock.Setup(r => r.ObtenerUltimoTurnoAsync(
                It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<ModoConsultaTurno>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildTurnoExistente(codigoTurno, cerrado: false));
        mock.Setup(r => r.ReAperturarAsync(
                It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<decimal>(), It.IsAny<decimal>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);
        return mock.Object;
    }

    private static AbrirTurnoHandler BuildHandler(
        ITurnoRepository turnoRepo,
        ITipoCambioRepository? tcRepo = null)
        => new(turnoRepo, DiaContableService(), tcRepo);

    // ─────────────────────────────────────────────────────────────────────────
    // BR-TURNO-001: Apertura nueva
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task AperturaNueva_SinTurnoExistente_Inserta()
    {
        var turnoRepo = TurnoRepoSinTurnoExistente("2600000001");
        var handler = BuildHandler(turnoRepo);
        var cmd = new AbrirTurnoCommand("CA1", "USR1", MontoInicial: 100m);

        var result = await handler.HandleAsync(cmd);

        Assert.True(result.EsExitoso);
        Assert.False(result.Valor!.EsReApertura);
        Assert.Equal("2600000001", result.Valor.CodigoTurno);
    }

    [Fact]
    public async Task AperturaNueva_ConMontoME_InsertaMontoME()
    {
        Domain.Entities.Configuracion.Turno? capturado = null;
        var mock = new Mock<ITurnoRepository>();
        mock.Setup(r => r.ObtenerUltimoTurnoAsync(
                It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<ModoConsultaTurno>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((TurnoExistente?)null);
        mock.Setup(r => r.GenerarCorrelativoAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync("2600000001");
        mock.Setup(r => r.InsertarAsync(It.IsAny<Domain.Entities.Configuracion.Turno>(), It.IsAny<CancellationToken>()))
            .Callback<Domain.Entities.Configuracion.Turno, CancellationToken>((t, _) => capturado = t)
            .ReturnsAsync(true);

        var handler = BuildHandler(mock.Object);
        var cmd = new AbrirTurnoCommand("CA1", "USR1", MontoInicial: 100m, MontoInicialME: 50m);

        var result = await handler.HandleAsync(cmd);

        Assert.True(result.EsExitoso);
        Assert.Equal(50m, capturado!.MontoInicialME);
    }

    [Fact]
    public async Task AperturaNueva_InsertFalla_RetornaError()
    {
        var mock = new Mock<ITurnoRepository>();
        mock.Setup(r => r.ObtenerUltimoTurnoAsync(
                It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<ModoConsultaTurno>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((TurnoExistente?)null);
        mock.Setup(r => r.GenerarCorrelativoAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync("2600000001");
        mock.Setup(r => r.InsertarAsync(It.IsAny<Domain.Entities.Configuracion.Turno>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(false);

        var handler = BuildHandler(mock.Object);
        var result = await handler.HandleAsync(new AbrirTurnoCommand("CA1", "USR1", 100m));

        Assert.False(result.EsExitoso);
        Assert.Equal("TURNO_APERTURA_FALLIDA", result.CodigoError);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // BR-TURNO-002: Re-apertura
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task ReApertura_TurnoAbierto_ActualizaMonto()
    {
        var turnoRepo = TurnoRepoConTurnoAbierto("2600000001");
        var handler = BuildHandler(turnoRepo);
        var cmd = new AbrirTurnoCommand("CA1", "USR1", MontoInicial: 200m);

        var result = await handler.HandleAsync(cmd);

        Assert.True(result.EsExitoso);
        Assert.True(result.Valor!.EsReApertura);
        Assert.Equal("2600000001", result.Valor.CodigoTurno);
    }

    [Fact]
    public async Task ReApertura_UpdateFalla_RetornaError()
    {
        var mock = new Mock<ITurnoRepository>();
        mock.Setup(r => r.ObtenerUltimoTurnoAsync(
                It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<ModoConsultaTurno>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildTurnoExistente(cerrado: false));
        mock.Setup(r => r.ReAperturarAsync(
                It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<decimal>(), It.IsAny<decimal>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(false);

        var handler = BuildHandler(mock.Object);
        var result = await handler.HandleAsync(new AbrirTurnoCommand("CA1", "USR1", 100m));

        Assert.False(result.EsExitoso);
        Assert.Equal("TURNO_REAPERTURA_FALLIDA", result.CodigoError);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // BR-TC-001: Tipo de cambio
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task RegistrarTipoCambio_SiIndicado_LlamaInsertarOActualizar()
    {
        var turnoRepo = TurnoRepoSinTurnoExistente();
        var tcMock = new Mock<ITipoCambioRepository>();
        tcMock.Setup(r => r.InsertarOActualizarAsync(
                It.IsAny<TipoCambioDelDia>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Returns(Task.FromResult(true));

        var handler = BuildHandler(turnoRepo, tcMock.Object);
        var cmd = new AbrirTurnoCommand("CA1", "USR1", 100m,
            RegistrarTipoCambio: true,
            TipoCambio: 3.75m, TipoCambioOficial: 3.70m);

        await handler.HandleAsync(cmd);

        tcMock.Verify(r => r.InsertarOActualizarAsync(
            It.Is<TipoCambioDelDia>(t => t.Venta == 3.75m && t.Oficial == 3.70m),
            "USR1",
            It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task RegistrarTipoCambio_NoIndicado_NoLlamaTC()
    {
        var turnoRepo = TurnoRepoSinTurnoExistente();
        var tcMock = new Mock<ITipoCambioRepository>();

        var handler = BuildHandler(turnoRepo, tcMock.Object);
        await handler.HandleAsync(new AbrirTurnoCommand("CA1", "USR1", 100m, RegistrarTipoCambio: false));

        tcMock.Verify(r => r.InsertarOActualizarAsync(
            It.IsAny<TipoCambioDelDia>(), It.IsAny<string>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Validaciones básicas
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task CajaVacia_RetornaError()
    {
        var handler = BuildHandler(TurnoRepoSinTurnoExistente());
        var result = await handler.HandleAsync(new AbrirTurnoCommand("", "USR1", 100m));

        Assert.False(result.EsExitoso);
        Assert.Equal("TURNO_CAJA_REQUERIDA", result.CodigoError);
    }

    [Fact]
    public async Task UsuarioVacio_RetornaError()
    {
        var handler = BuildHandler(TurnoRepoSinTurnoExistente());
        var result = await handler.HandleAsync(new AbrirTurnoCommand("CA1", "", 100m));

        Assert.False(result.EsExitoso);
        Assert.Equal("TURNO_USUARIO_REQUERIDO", result.CodigoError);
    }
}
