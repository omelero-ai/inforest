using Inforest.Application.Caja;
using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Entities.Ventas;
using Moq;

namespace Inforest.Application.Tests.Caja;

/// <summary>
/// Tests para RegistrarPagosMultiplesHandler.
/// Legacy: frmPago.frm — cmdOpcion_Click(2) "Grabar".
/// BR-PAGO-001: DPAGODOCUMENTO acepta múltiples filas por documento.
/// BR-PAGO-002: MDOCUMENTO se marca cobrado cuando la suma de pagos cubre el total.
/// </summary>
public class RegistrarPagosMultiplesHandlerTests
{
    private static Documento BuildDocumento(string codigo = "DOC001", decimal neto = 100m, decimal imp1 = 18m)
        => Documento.Emitir(codigo, "01", null, neto, imp1, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");

    private static RegistrarPagosMultiplesHandler BuildHandler(
        Mock<IPagoRepository> pagoRepo,
        Mock<IDocumentoRepository> docRepo)
    {
        var parametroRepo = new Mock<IParametroRepository>();
        parametroRepo.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync((Inforest.Domain.Entities.Configuracion.ConfiguracionSistema?)null);

        var impresora = new Mock<IImpresoraService>();
        impresora.Setup(i => i.ImprimirTicketAsync(It.IsAny<string>(), It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        return new RegistrarPagosMultiplesHandler(
            pagoRepo.Object,
            docRepo.Object,
            parametroRepo.Object,
            impresora.Object);
    }

    // BR-PAGO-001: insertar una fila por cada medio de pago
    [Fact]
    public async Task MultiPago_DosLineas_InsertaDosPagosYMarcaCobrado()
    {
        var documento = BuildDocumento();
        var pagoRepo = new Mock<IPagoRepository>();
        pagoRepo.Setup(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);
        docRepo.Setup(r => r.MarcarCobradoAsync("DOC001", 118m, It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = BuildHandler(pagoRepo, docRepo);
        var pagos = new List<LineaPago>
        {
            new("01", 100m, 0m, null),   // efectivo MN
            new("02", 18m, 0m, "TRX123") // tarjeta
        };

        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC001", pagos));

        Assert.True(result.EsExitoso);
        // BR-PAGO-001: dos filas insertadas en DPAGODOCUMENTO
        pagoRepo.Verify(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()), Times.Exactly(2));
        // BR-PAGO-002: documento marcado cobrado con total 118
        docRepo.Verify(r => r.MarcarCobradoAsync("DOC001", 118m, It.IsAny<CancellationToken>()), Times.Once);
    }

    // BR-PAGO-002: suma de pagos debe cubrir el total
    [Fact]
    public async Task MultiPago_SumaInsuficiente_RetornaError()
    {
        var documento = BuildDocumento(); // total 118m
        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);

        var pagoRepo = new Mock<IPagoRepository>();
        var handler = BuildHandler(pagoRepo, docRepo);

        var pagos = new List<LineaPago> { new("01", 50m, 0m, null) }; // 50 < 118
        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC001", pagos));

        Assert.False(result.EsExitoso);
        Assert.Equal("CAJA_PAGO_INSUFICIENTE", result.CodigoError);
        pagoRepo.Verify(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task MultiPago_ListaVacia_RetornaError()
    {
        var docRepo = new Mock<IDocumentoRepository>();
        var pagoRepo = new Mock<IPagoRepository>();
        var handler = BuildHandler(pagoRepo, docRepo);

        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC001", []));

        Assert.False(result.EsExitoso);
        Assert.Equal("CAJA_PAGOS_VACIOS", result.CodigoError);
    }

    [Fact]
    public async Task MultiPago_DocumentoNoEncontrado_RetornaError()
    {
        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC_NOPE", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Documento?)null);

        var pagoRepo = new Mock<IPagoRepository>();
        var handler = BuildHandler(pagoRepo, docRepo);

        var pagos = new List<LineaPago> { new("01", 200m, 0m, null) };
        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC_NOPE", pagos));

        Assert.False(result.EsExitoso);
        Assert.Equal("CAJA_DOCUMENTO_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task MultiPago_DocumentoYaCobrado_RetornaErrorDominio()
    {
        var documento = BuildDocumento();
        documento.Cobrar(118m); // ya cobrado

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);

        var pagoRepo = new Mock<IPagoRepository>();
        var handler = BuildHandler(pagoRepo, docRepo);

        var pagos = new List<LineaPago> { new("01", 150m, 0m, null) };
        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC001", pagos));

        Assert.False(result.EsExitoso);
        Assert.Equal("DOCUMENTO_YA_COBRADO", result.CodigoError);
    }

    [Fact]
    public async Task MultiPago_DocumentoAnulado_RetornaErrorDominio()
    {
        var documento = BuildDocumento();
        documento.Anular("USR01", "Test");

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);

        var pagoRepo = new Mock<IPagoRepository>();
        var handler = BuildHandler(pagoRepo, docRepo);

        var pagos = new List<LineaPago> { new("01", 150m, 0m, null) };
        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC001", pagos));

        Assert.False(result.EsExitoso);
        Assert.Equal("DOCUMENTO_ANULADO", result.CodigoError);
    }

    // Pago exacto (sin vuelto)
    [Fact]
    public async Task MultiPago_MontoExacto_Exitoso()
    {
        var documento = BuildDocumento(); // total = 118m
        var pagoRepo = new Mock<IPagoRepository>();
        pagoRepo.Setup(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);
        docRepo.Setup(r => r.MarcarCobradoAsync("DOC001", 118m, It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = BuildHandler(pagoRepo, docRepo);
        var pagos = new List<LineaPago> { new("01", 118m, 0m, null) };

        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC001", pagos));

        Assert.True(result.EsExitoso);
    }

    // Pago con vuelto (monto mayor al total)
    [Fact]
    public async Task MultiPago_MontoConVuelto_Exitoso()
    {
        var documento = BuildDocumento(); // total = 118m
        var pagoRepo = new Mock<IPagoRepository>();
        pagoRepo.Setup(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);
        docRepo.Setup(r => r.MarcarCobradoAsync("DOC001", 200m, It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = BuildHandler(pagoRepo, docRepo);
        var pagos = new List<LineaPago> { new("01", 200m, 82m, null) }; // 200 - 118 = 82 de vuelto

        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC001", pagos));

        Assert.True(result.EsExitoso);
    }

    // BR-PAGO-001: un solo medio de pago también es válido
    [Fact]
    public async Task MultiPago_UnSoloMedio_InserataUnPago()
    {
        var documento = BuildDocumento();
        var pagoRepo = new Mock<IPagoRepository>();
        pagoRepo.Setup(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);
        docRepo.Setup(r => r.MarcarCobradoAsync("DOC001", 120m, It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = BuildHandler(pagoRepo, docRepo);
        var pagos = new List<LineaPago> { new("01", 120m, 2m, null) };

        var result = await handler.HandleAsync(new RegistrarPagosMultiplesCommand("DOC001", pagos));

        Assert.True(result.EsExitoso);
        pagoRepo.Verify(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()), Times.Once);
    }
}
