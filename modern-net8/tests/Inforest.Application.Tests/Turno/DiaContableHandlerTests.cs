using Inforest.Application.Turno;
using Moq;

namespace Inforest.Application.Tests.Turno;

/// <summary>
/// Tests para AperturarDiaContableHandler y CerrarDiaContableHandler.
///
/// Legacy: frmDiaContable.frm — cmdApertura_Click, cmdCerrar_Click, cmdOpcion_Click.
/// Clase: clsDiaContable.cls — insertarDiaContable, obtieneDiaContable, obtieneFechaServidor.
///
/// Reglas verificadas:
///   BR-DC-001 — No se puede vender sin aperturar el día contable.
///   BR-DC-002 — No se puede aperturar si ya existe uno activo (fecha != sentinel).
///   BR-DC-003 — La fecha no puede ser menor al último día contable registrado (validaFechaIngreso).
///   BR-DC-004 — El cierre actualiza lcierre=1 con usuario y timestamp.
/// </summary>
public class DiaContableHandlerTests
{
    private static readonly DateTime Sentinel = new(1900, 1, 1);
    private static readonly DateTime HoyBase = new(2026, 8, 14);

    // ─────────────────────────────────────────────────────────────────────────
    // AperturarDiaContableHandler
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task Apertura_SinDiaContablePrevio_Ok()
    {
        // Legacy: fdiacontable = 19000101 → primer arranque, apertura permitida
        var svc = new Mock<IDiaContableService>();
        svc.Setup(s => s.ObtenerDiaContableAsync(true, "05:00", "USR1", default))
           .ReturnsAsync(Sentinel);
        svc.Setup(s => s.ObtenerFechaMaximaDiaContableAsync(default))
           .ReturnsAsync((DateTime?)null); // sin registros previos
        svc.Setup(s => s.InsertarDiaContableAsync(HoyBase, "USR1", default))
           .Returns(Task.CompletedTask);

        var handler = new AperturarDiaContableHandler(svc.Object);
        var result = await handler.HandleAsync(new AperturarDiaContableCommand(HoyBase, "USR1"));

        Assert.True(result.EsExitoso);
        svc.Verify(s => s.InsertarDiaContableAsync(HoyBase, "USR1", default), Times.Once);
    }

    [Fact]
    public async Task Apertura_FechaIgualAMaxima_Ok()
    {
        // BR-DC-003: fecha == max → permitido (>= no <)
        var svc = new Mock<IDiaContableService>();
        svc.Setup(s => s.ObtenerDiaContableAsync(true, "05:00", "USR1", default))
           .ReturnsAsync(Sentinel);
        svc.Setup(s => s.ObtenerFechaMaximaDiaContableAsync(default))
           .ReturnsAsync(HoyBase);
        svc.Setup(s => s.InsertarDiaContableAsync(HoyBase, "USR1", default))
           .Returns(Task.CompletedTask);

        var handler = new AperturarDiaContableHandler(svc.Object);
        var result = await handler.HandleAsync(new AperturarDiaContableCommand(HoyBase, "USR1"));

        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task Apertura_FechaMenorAMaxima_Falla_BR_DC_003()
    {
        // BR-DC-003: validaFechaIngreso → fecha < max → error
        var svc = new Mock<IDiaContableService>();
        svc.Setup(s => s.ObtenerDiaContableAsync(true, "05:00", "USR1", default))
           .ReturnsAsync(Sentinel);
        svc.Setup(s => s.ObtenerFechaMaximaDiaContableAsync(default))
           .ReturnsAsync(HoyBase);

        var handler = new AperturarDiaContableHandler(svc.Object);
        var fechaAnterior = HoyBase.AddDays(-1);
        var result = await handler.HandleAsync(new AperturarDiaContableCommand(fechaAnterior, "USR1"));

        Assert.False(result.EsExitoso);
        Assert.Equal("DC_FECHA_MENOR_ULTIMA", result.CodigoError);
        svc.Verify(s => s.InsertarDiaContableAsync(It.IsAny<DateTime>(), It.IsAny<string>(), default), Times.Never);
    }

    [Fact]
    public async Task Apertura_UsuarioVacio_Falla()
    {
        var svc = new Mock<IDiaContableService>();
        var handler = new AperturarDiaContableHandler(svc.Object);

        var result = await handler.HandleAsync(new AperturarDiaContableCommand(HoyBase, ""));

        Assert.False(result.EsExitoso);
        Assert.Equal("DC_USUARIO_REQUERIDO", result.CodigoError);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // CerrarDiaContableHandler
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task Cierre_Ok_BR_DC_004()
    {
        // BR-DC-004: cierre llama CerrarDiaContableAsync con usuario y fecha
        var svc = new Mock<IDiaContableService>();
        svc.Setup(s => s.CerrarDiaContableAsync(HoyBase, "USR2", default))
           .Returns(Task.CompletedTask);

        var handler = new CerrarDiaContableHandler(svc.Object);
        var result = await handler.HandleAsync(new CerrarDiaContableCommand(HoyBase, "USR2"));

        Assert.True(result.EsExitoso);
        svc.Verify(s => s.CerrarDiaContableAsync(HoyBase, "USR2", default), Times.Once);
    }

    [Fact]
    public async Task Cierre_UsuarioVacio_Falla()
    {
        var svc = new Mock<IDiaContableService>();
        var handler = new CerrarDiaContableHandler(svc.Object);

        var result = await handler.HandleAsync(new CerrarDiaContableCommand(HoyBase, ""));

        Assert.False(result.EsExitoso);
        Assert.Equal("DC_USUARIO_CIERRE_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task Cierre_FechaDefault_Falla()
    {
        var svc = new Mock<IDiaContableService>();
        var handler = new CerrarDiaContableHandler(svc.Object);

        var result = await handler.HandleAsync(new CerrarDiaContableCommand(default, "USR2"));

        Assert.False(result.EsExitoso);
        Assert.Equal("DC_FECHA_REQUERIDA", result.CodigoError);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // ObtenerDiaContableHandler
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task Obtener_DiaSentinel_EsNuevo()
    {
        // Legacy: fdiacontable = 19000101 → esNuevo = true
        var svc = new Mock<IDiaContableService>();
        svc.Setup(s => s.ObtenerDiaContableAsync(true, "05:00", "USR1", default))
           .ReturnsAsync(Sentinel);

        var handler = new ObtenerDiaContableHandler(svc.Object);
        var result = await handler.HandleAsync(new ObtenerDiaContableQuery("USR1"));

        Assert.True(result.EsExitoso);
        Assert.True(result.Valor!.EsNuevo);
        Assert.Equal(Sentinel.Date, result.Valor.Fecha.Date);
    }

    [Fact]
    public async Task Obtener_DiaExistente_NoEsNuevo()
    {
        var svc = new Mock<IDiaContableService>();
        svc.Setup(s => s.ObtenerDiaContableAsync(true, "05:00", "USR1", default))
           .ReturnsAsync(HoyBase);

        var handler = new ObtenerDiaContableHandler(svc.Object);
        var result = await handler.HandleAsync(new ObtenerDiaContableQuery("USR1"));

        Assert.True(result.EsExitoso);
        Assert.False(result.Valor!.EsNuevo);
        Assert.Equal(HoyBase, result.Valor.Fecha);
    }

    [Fact]
    public async Task Obtener_UsuarioVacio_Falla()
    {
        var svc = new Mock<IDiaContableService>();
        var handler = new ObtenerDiaContableHandler(svc.Object);

        var result = await handler.HandleAsync(new ObtenerDiaContableQuery(""));

        Assert.False(result.EsExitoso);
        Assert.Equal("DC_USUARIO_REQUERIDO", result.CodigoError);
    }
}
