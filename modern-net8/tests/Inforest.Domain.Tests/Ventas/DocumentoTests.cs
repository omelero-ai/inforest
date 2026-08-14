using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Ventas;

public class DocumentoTests
{
    [Fact]
    public void Emitir_DocumentoValido_QuedaPendienteYCalculaTotal()
    {
        var documento = Documento.Emitir("B001-0001", "01", "CLI01", 100m, 18m, 0m, 0m, 10m, 5m, "TUR001", "CAJ01");

        Assert.Equal("B001-0001", documento.CodigoDocumento);
        Assert.Equal("PE", documento.Estado);
        Assert.Equal(123m, documento.Total);
    }

    [Fact]
    public void Cobrar_DocumentoYaCobrado_LanzaExcepcion()
    {
        var documento = Documento.Emitir("B001-0002", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");
        documento.Cobrar(118m);

        var ex = Assert.Throws<DomainException>(() => documento.Cobrar(118m));

        Assert.Equal("DOCUMENTO_YA_COBRADO", ex.Code);
    }

    [Fact]
    public void CalcularTotal_ConImpuestosYDescuento_RetornaValorEsperado()
    {
        var documento = Documento.Emitir("F001-0001", "03", "CLI02", 200m, 36m, 4m, 0m, 20m, 10m, "TUR002", "CAJ02");

        Assert.Equal(250m, documento.CalcularTotal());
    }
}
