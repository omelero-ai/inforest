using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Seguridad;
using Moq;

namespace Inforest.Application.Tests.Ventas;

/// <summary>
/// Tests del handler CambiarDocumentoHandler.
/// BR-CAMBIO-001..005.
/// Legacy: frmCambio.frm.
/// </summary>
public class CambiarDocumentoHandlerTests
{
    private static SesionOperativa BuildSesion(
        string caja = "CAJ01", string turno = "TUR01", string salon = "SAL01")
        => SesionOperativa.Iniciar(
            "USR01", "Usuario", "GRP01", "INFOREST",
            caja, "TERM01", "INFOREST", 1, DateTime.Now, []);

    private static CambiarDocumentoCommand BuildCommand(
        string origen = "B-00001-000000042",
        string nuevo = "F-00001-000000001",
        string tipoEmision = "01",
        string cliente = "10450614767",
        string motivo = "Cliente solicita factura",
        bool esConsumo = false)
        => new(origen, tipoEmision, nuevo, cliente, motivo, esConsumo, string.Empty);

    // ── BR-CAMBIO-003: Motivo requerido ──────────────────────────────────────

    [Fact]
    public async Task Cambiar_SinMotivo_RetornaError()
    {
        var repo = new Mock<ICambioDocumentoRepository>();
        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new CambiarDocumentoHandler(repo.Object, session.Object);
        var result = await handler.HandleAsync(BuildCommand(motivo: ""));

        Assert.False(result.EsExitoso);
        Assert.Equal("CAMBIO_MOTIVO_REQUERIDO", result.CodigoError);
    }

    // ── Sin sesión ────────────────────────────────────────────────────────────

    [Fact]
    public async Task Cambiar_SinSesion_RetornaError()
    {
        var repo = new Mock<ICambioDocumentoRepository>();
        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns((SesionOperativa?)null);

        var handler = new CambiarDocumentoHandler(repo.Object, session.Object);
        var result = await handler.HandleAsync(BuildCommand());

        Assert.False(result.EsExitoso);
        Assert.Equal("CAMBIO_SIN_SESION", result.CodigoError);
    }

    // ── BR-CAMBIO-002: Monto excede máximo ───────────────────────────────────

    [Fact]
    public async Task Cambiar_MontoExcedeMaximo_RetornaError()
    {
        var repo = new Mock<ICambioDocumentoRepository>();
        repo.Setup(r => r.ObtenerMontosValidacionAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((700m, 500m)); // doc=700, max=500

        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new CambiarDocumentoHandler(repo.Object, session.Object);
        var result = await handler.HandleAsync(BuildCommand());

        Assert.False(result.EsExitoso);
        Assert.Equal("CAMBIO_MONTO_EXCEDE_MAXIMO", result.CodigoError);
    }

    // ── BR-CAMBIO-002: Monto igual al máximo debe pasar ──────────────────────

    [Fact]
    public async Task Cambiar_MontoIgualMaximo_Ejecuta()
    {
        var repo = new Mock<ICambioDocumentoRepository>();
        repo.Setup(r => r.ObtenerMontosValidacionAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((500m, 500m)); // doc=500, max=500 — debe pasar (condición < no <=)
        repo.Setup(r => r.EjecutarCambioAsync(It.IsAny<EjecutarCambioDocumentoParams>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        repo.Setup(r => r.ActualizarEmisionAsync(It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new CambiarDocumentoHandler(repo.Object, session.Object);
        var result = await handler.HandleAsync(BuildCommand());

        Assert.True(result.EsExitoso);
    }

    // ── Sin máximo (nMontoMaximo=0) no valida monto ───────────────────────────

    [Fact]
    public async Task Cambiar_MaximoCero_NoValidaMonto_Ejecuta()
    {
        var repo = new Mock<ICambioDocumentoRepository>();
        repo.Setup(r => r.ObtenerMontosValidacionAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((9999m, 0m)); // máximo=0 → sin límite
        repo.Setup(r => r.EjecutarCambioAsync(It.IsAny<EjecutarCambioDocumentoParams>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        repo.Setup(r => r.ActualizarEmisionAsync(It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new CambiarDocumentoHandler(repo.Object, session.Object);
        var result = await handler.HandleAsync(BuildCommand());

        Assert.True(result.EsExitoso);
    }

    // ── Flujo exitoso: retorna CodigoDocumentoNuevo ───────────────────────────

    [Fact]
    public async Task Cambiar_Exitoso_RetornaCodigoNuevo()
    {
        var repo = new Mock<ICambioDocumentoRepository>();
        repo.Setup(r => r.ObtenerMontosValidacionAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((200m, 0m));
        repo.Setup(r => r.EjecutarCambioAsync(It.IsAny<EjecutarCambioDocumentoParams>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        repo.Setup(r => r.ActualizarEmisionAsync(It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new CambiarDocumentoHandler(repo.Object, session.Object);
        var result = await handler.HandleAsync(BuildCommand(nuevo: "F-00001-000000001"));

        Assert.True(result.EsExitoso);
        Assert.Equal("F-00001-000000001", result.Valor!.CodigoDocumentoNuevo);
    }

    // ── Falla en el repo ──────────────────────────────────────────────────────

    [Fact]
    public async Task Cambiar_ExcepcionRepo_RetornaError()
    {
        var repo = new Mock<ICambioDocumentoRepository>();
        repo.Setup(r => r.ObtenerMontosValidacionAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("DB error"));

        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new CambiarDocumentoHandler(repo.Object, session.Object);
        var result = await handler.HandleAsync(BuildCommand());

        Assert.False(result.EsExitoso);
        Assert.Equal("CAMBIO_ERROR", result.CodigoError);
    }

    // ── Código documento nuevo requerido ─────────────────────────────────────

    [Fact]
    public async Task Cambiar_SinDocumentoNuevo_RetornaError()
    {
        var repo = new Mock<ICambioDocumentoRepository>();
        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new CambiarDocumentoHandler(repo.Object, session.Object);
        var result = await handler.HandleAsync(BuildCommand(nuevo: ""));

        Assert.False(result.EsExitoso);
        Assert.Equal("CAMBIO_NUEVO_REQUERIDO", result.CodigoError);
    }
}
