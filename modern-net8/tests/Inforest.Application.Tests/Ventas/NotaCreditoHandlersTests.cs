using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;
using Moq;

namespace Inforest.Application.Tests.Ventas;

/// <summary>
/// Tests de los handlers de Nota de Crédito.
/// BR-NC-001..006, BR-NC-003.
/// </summary>
public class NotaCreditoHandlersTests
{
    private static Domain.Entities.Seguridad.SesionOperativa BuildSesion(string codigoCaja = "CAJ01")
        => Domain.Entities.Seguridad.SesionOperativa.Iniciar(
            "USR01", "Usuario Prueba", "GRP01", "INFOREST",
            codigoCaja, "TERM01", "INFOREST", 1, DateTime.Now, []);

    private static Documento BuildDocumento(string codigo = "DOC001", decimal total = 500m)
        => Documento.Reconstituir(
            codigo, "01", null,
            total / 1.18m, total - (total / 1.18m), 0m, 0m,
            total, 0m, 0m, total,
            "PE", "TUR01", "CAJ01",
            DateTime.Now.AddHours(-1), null, null);

    // ── EmitirNotaCreditoHandler ──────────────────────────────────────────────

    [Fact]
    public async Task EmitirNC_DocumentoValido_RetornaNC()
    {
        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildDocumento());

        var ncRepo = new Mock<INotaCreditoRepository>();
        ncRepo.Setup(r => r.ObtenerTotalEmitidoParaDocumentoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(0m);
        ncRepo.Setup(r => r.InsertarAsync(It.IsAny<NotaCredito>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new EmitirNotaCreditoHandler(ncRepo.Object, docRepo.Object, session.Object);
        var result = await handler.HandleAsync(new EmitirNotaCreditoCommand(
            "NC001", DateTime.Today, "DOC001", 100m, 18m, 0m, 0m, "Devolución"));

        Assert.True(result.EsExitoso);
        Assert.Equal("NC001", result.Valor!.CodigoNotaCredito);
        Assert.Equal("PE", result.Valor!.Estado);
    }

    [Fact]
    public async Task EmitirNC_SinSesion_RetornaError()
    {
        var docRepo = new Mock<IDocumentoRepository>();
        var ncRepo = new Mock<INotaCreditoRepository>();
        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns((Domain.Entities.Seguridad.SesionOperativa?)null);

        var handler = new EmitirNotaCreditoHandler(ncRepo.Object, docRepo.Object, session.Object);
        var result = await handler.HandleAsync(new EmitirNotaCreditoCommand(
            "NC002", DateTime.Today, "DOC001", 50m, 9m, 0m, 0m, "Motivo"));

        Assert.False(result.EsExitoso);
        Assert.Equal("NC_SIN_SESION", result.CodigoError);
    }

    [Fact]
    public async Task EmitirNC_DocumentoNoExiste_RetornaError()
    {
        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOCXXX", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Documento?)null);

        var ncRepo = new Mock<INotaCreditoRepository>();
        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new EmitirNotaCreditoHandler(ncRepo.Object, docRepo.Object, session.Object);
        var result = await handler.HandleAsync(new EmitirNotaCreditoCommand(
            "NC003", DateTime.Today, "DOCXXX", 50m, 9m, 0m, 0m, "Motivo"));

        Assert.False(result.EsExitoso);
        Assert.Equal("NC_DOCUMENTO_NO_ENCONTRADO", result.CodigoError);
    }

    // BR-NC-003: NC total no puede superar saldo disponible del documento
    [Fact]
    public async Task EmitirNC_ExcedeSaldoDocumento_RetornaError()
    {
        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildDocumento(total: 100m));

        var ncRepo = new Mock<INotaCreditoRepository>();
        // ya hay 80 en NCs previas, saldo = 20
        ncRepo.Setup(r => r.ObtenerTotalEmitidoParaDocumentoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(80m);
        // intentamos emitir NC por 50 → 80+50 > 100
        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new EmitirNotaCreditoHandler(ncRepo.Object, docRepo.Object, session.Object);
        var result = await handler.HandleAsync(new EmitirNotaCreditoCommand(
            "NC004", DateTime.Today, "DOC001", 50m, 0m, 0m, 0m, "Excede"));

        Assert.False(result.EsExitoso);
        Assert.Equal("NC_EXCEDE_DOCUMENTO", result.CodigoError);
    }

    [Fact]
    public async Task EmitirNC_DocumentoAnulado_RetornaError()
    {
        var docAnulado = Documento.Reconstituir(
            "DOC_AN", "01", null, 100m, 18m, 0m, 0m, 118m, 0m, 0m, 118m,
            "AN", "TUR01", "CAJ01", DateTime.Now, null, null);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC_AN", It.IsAny<CancellationToken>()))
            .ReturnsAsync(docAnulado);

        var ncRepo = new Mock<INotaCreditoRepository>();
        var session = new Mock<ISessionService>();
        session.Setup(s => s.SesionActual).Returns(BuildSesion());

        var handler = new EmitirNotaCreditoHandler(ncRepo.Object, docRepo.Object, session.Object);
        var result = await handler.HandleAsync(new EmitirNotaCreditoCommand(
            "NC005", DateTime.Today, "DOC_AN", 50m, 0m, 0m, 0m, "Motivo"));

        Assert.False(result.EsExitoso);
        Assert.Equal("NC_DOCUMENTO_ANULADO", result.CodigoError);
    }

    // ── AnularNotaCreditoHandler ──────────────────────────────────────────────

    [Fact]
    public async Task AnularNC_Exitosa_RetornaOk()
    {
        var nc = NotaCredito.Emitir("NC006", DateTime.Today, "DOC001", 50m, 9m, 0m, 0m, "Motivo");
        var ncRepo = new Mock<INotaCreditoRepository>();
        ncRepo.Setup(r => r.ObtenerPorCodigoAsync("NC006", It.IsAny<CancellationToken>()))
            .ReturnsAsync(nc);
        ncRepo.Setup(r => r.AnularAsync("NC006", "USR01", It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = new AnularNotaCreditoHandler(ncRepo.Object);
        var result = await handler.HandleAsync(new AnularNotaCreditoCommand("NC006", "USR01"));

        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task AnularNC_NoEncontrada_RetornaError()
    {
        var ncRepo = new Mock<INotaCreditoRepository>();
        ncRepo.Setup(r => r.ObtenerPorCodigoAsync("NCXXX", It.IsAny<CancellationToken>()))
            .ReturnsAsync((NotaCredito?)null);

        var handler = new AnularNotaCreditoHandler(ncRepo.Object);
        var result = await handler.HandleAsync(new AnularNotaCreditoCommand("NCXXX", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("NC_NO_ENCONTRADA", result.CodigoError);
    }

    // ── ObtenerNotasCreditoPorFechaHandler ────────────────────────────────────

    [Fact]
    public async Task ObtenerPorFecha_RangoInvalido_RetornaError()
    {
        var ncRepo = new Mock<INotaCreditoRepository>();
        var handler = new ObtenerNotasCreditoPorFechaHandler(ncRepo.Object);

        var result = await handler.HandleAsync(
            new ObtenerNotasCreditoPorFechaQuery(DateTime.Today, DateTime.Today.AddDays(-1)));

        Assert.False(result.EsExitoso);
        Assert.Equal("NC_RANGO_FECHAS_INVALIDO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerPorFecha_RangoValido_RetornaLista()
    {
        var ncRepo = new Mock<INotaCreditoRepository>();
        ncRepo.Setup(r => r.ObtenerPorRangoFechaAsync(It.IsAny<DateTime>(), It.IsAny<DateTime>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<NotaCredito>().AsReadOnly());

        var handler = new ObtenerNotasCreditoPorFechaHandler(ncRepo.Object);
        var result = await handler.HandleAsync(
            new ObtenerNotasCreditoPorFechaQuery(DateTime.Today, DateTime.Today));

        Assert.True(result.EsExitoso);
        Assert.Empty(result.Valor!);
    }
}
