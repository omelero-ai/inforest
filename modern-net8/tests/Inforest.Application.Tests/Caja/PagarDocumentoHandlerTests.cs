using Inforest.Application.Caja;
using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Entities.Ventas;
using Moq;

namespace Inforest.Application.Tests.Caja;

public class PagarDocumentoHandlerTests
{
    [Fact]
    public async Task CobroExitoso_LlamaPagoRepoYDocumentoRepo()
    {
        var documento = Documento.Emitir("DOC001", "01", "CLI01", 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");

        var pagoRepo = new Mock<IPagoRepository>();
        pagoRepo.Setup(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var documentoRepo = new Mock<IDocumentoRepository>();
        documentoRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);
        documentoRepo.Setup(r => r.MarcarCobradoAsync("DOC001", 118m, It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var parametroRepo = new Mock<IParametroRepository>();
        parametroRepo.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync((Inforest.Domain.Entities.Configuracion.ConfiguracionSistema?)null);

        var impresora = new Mock<IImpresoraService>();
        impresora.Setup(i => i.ImprimirTicketAsync(It.IsAny<string>(), It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var handler = new PagarDocumentoHandler(
            pagoRepo.Object,
            documentoRepo.Object,
            parametroRepo.Object,
            impresora.Object);

        var result = await handler.HandleAsync(new PagarDocumentoCommand("DOC001", "01", 118m, 0m, "REF1"));

        Assert.True(result.EsExitoso);
        pagoRepo.Verify(r => r.InsertarPagoAsync(It.IsAny<PagoDocumento>(), It.IsAny<CancellationToken>()), Times.Once);
        documentoRepo.Verify(r => r.MarcarCobradoAsync("DOC001", 118m, It.IsAny<CancellationToken>()), Times.Once);
    }
}
