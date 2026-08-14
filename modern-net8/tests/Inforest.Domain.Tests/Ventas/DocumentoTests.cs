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

    // BR-VENTA-001: estado inicial PE (pendiente de pago)
    [Fact]
    public void Emitir_EstadoInicialEsPendiente()
    {
        var documento = Documento.Emitir("DOC001", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");

        Assert.Equal("PE", documento.Estado);
        Assert.Null(documento.FechaPago);
    }

    // Fórmula: total = neto + imp1 + imp2 + imp3 + propina - descuento
    [Fact]
    public void Emitir_PropinaSumaAlTotal()
    {
        var doc = Documento.Emitir("DOC002", "01", null, 100m, 18m, 0m, 0m, 5m, 0m, "TUR001", "CAJ01");
        // 100 + 18 + 5 = 123
        Assert.Equal(123m, doc.Total);
    }

    [Fact]
    public void Emitir_DescuentoRestaDelTotal()
    {
        var doc = Documento.Emitir("DOC003", "01", null, 100m, 18m, 0m, 0m, 0m, 10m, "TUR001", "CAJ01");
        // 100 + 18 - 10 = 108
        Assert.Equal(108m, doc.Total);
    }

    // BR-VENTA-002: Cobrar transiciona PE → CO y registra abono/vuelto
    [Fact]
    public void Cobrar_DocumentoPendiente_TransicionaACobrado()
    {
        var doc = Documento.Emitir("DOC004", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");
        doc.Cobrar(120m);

        Assert.Equal("CO", doc.Estado);
        Assert.Equal(120m, doc.Abono);
        Assert.Equal(2m, doc.Vuelto);
        Assert.NotNull(doc.FechaPago);
    }

    // Cobrar con abono exacto: vuelto = 0
    [Fact]
    public void Cobrar_AbonoExacto_VueltoEsCero()
    {
        var doc = Documento.Emitir("DOC005", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");
        doc.Cobrar(118m);

        Assert.Equal(0m, doc.Vuelto);
        Assert.Equal("CO", doc.Estado);
    }

    // Cobrar con abono insuficiente retorna excepción
    [Fact]
    public void Cobrar_AbonoInsuficiente_LanzaExcepcion()
    {
        var doc = Documento.Emitir("DOC006", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");

        var ex = Assert.Throws<DomainException>(() => doc.Cobrar(50m));
        Assert.Equal("DOCUMENTO_ABONO_INSUFICIENTE", ex.Code);
    }

    // BR-VENTA-003: Anular transiciona PE → AN
    [Fact]
    public void Anular_DocumentoPendiente_TransicionaAAnulado()
    {
        var doc = Documento.Emitir("DOC007", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");
        doc.Anular("USR01", "Motivo de prueba");

        Assert.Equal("AN", doc.Estado);
        Assert.Equal("USR01", doc.UsuarioAnulado);
        Assert.Equal("Motivo de prueba", doc.MotivoAnulacion);
        Assert.NotNull(doc.FechaAnulacion);
    }

    // Documento cobrado puede anularse
    [Fact]
    public void Anular_DocumentoCobrado_TransicionaAAnulado()
    {
        var doc = Documento.Emitir("DOC008", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");
        doc.Cobrar(118m);
        doc.Anular("USR01", "Devolucion");

        Assert.Equal("AN", doc.Estado);
    }

    // No se puede cobrar un documento anulado
    [Fact]
    public void Cobrar_DocumentoAnulado_LanzaExcepcion()
    {
        var doc = Documento.Emitir("DOC009", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");
        doc.Anular("USR01", "Test");

        var ex = Assert.Throws<DomainException>(() => doc.Cobrar(118m));
        Assert.Equal("DOCUMENTO_ANULADO", ex.Code);
    }

    // Validaciones de campos obligatorios al emitir
    [Theory]
    [InlineData("", "01", "TUR001", "CAJ01", "DOCUMENTO_CODIGO_REQUERIDO")]
    [InlineData("DOC010", "", "TUR001", "CAJ01", "DOCUMENTO_TIPO_REQUERIDO")]
    [InlineData("DOC010", "01", "", "CAJ01", "DOCUMENTO_TURNO_REQUERIDO")]
    [InlineData("DOC010", "01", "TUR001", "", "DOCUMENTO_CAJA_REQUERIDA")]
    public void Emitir_CampoObligatorioVacio_LanzaExcepcion(string codigo, string tipo, string turno, string caja, string codigoError)
    {
        var ex = Assert.Throws<DomainException>(
            () => Documento.Emitir(codigo, tipo, null, 100m, 18m, 0m, 0m, 0m, 0m, turno, caja));

        Assert.Equal(codigoError, ex.Code);
    }
}
