using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Caja;

/// <summary>
/// Tests de dominio para ReciboIngreso.
/// Cubre BR-RECIBO-001, BR-RECIBO-002, BR-RECIBO-003, BR-RECIBO-004, BR-RECIBO-005, BR-RECIBO-011.
/// </summary>
public class ReciboIngresoTests
{
    [Fact]
    public void Registrar_ConDatosValidos_RetornaReciboEmitido()
    {
        var recibo = ReciboIngreso.Registrar("2600000001", DateTime.Today, 150m, "Ingreso de caja");

        Assert.Equal("2600000001", recibo.CodigoRecibo);
        Assert.Equal(150m, recibo.Monto);
        Assert.Equal("01", recibo.Estado);
        Assert.Equal("EMITIDO", recibo.DescripcionEstado);
        Assert.True(recibo.EstaEmitido);
    }

    [Fact]
    public void Registrar_MontoNegativo_LanzaExcepcion()
    {
        // BR-RECIBO-011
        var ex = Assert.Throws<DomainException>(() =>
            ReciboIngreso.Registrar("2600000001", DateTime.Today, -50m, "Descripcion"));

        Assert.Equal("RECIBO_MONTO_INVALIDO", ex.Code);
    }

    [Fact]
    public void Registrar_MontoZero_LanzaExcepcion()
    {
        // BR-RECIBO-011
        var ex = Assert.Throws<DomainException>(() =>
            ReciboIngreso.Registrar("2600000001", DateTime.Today, 0m, "Descripcion"));

        Assert.Equal("RECIBO_MONTO_INVALIDO", ex.Code);
    }

    [Fact]
    public void Registrar_DescripcionVacia_LanzaExcepcion()
    {
        // BR-RECIBO-011
        var ex = Assert.Throws<DomainException>(() =>
            ReciboIngreso.Registrar("2600000001", DateTime.Today, 100m, ""));

        Assert.Equal("RECIBO_DESCRIPCION_REQUERIDA", ex.Code);
    }

    [Fact]
    public void Registrar_CodigoVacio_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            ReciboIngreso.Registrar("", DateTime.Today, 100m, "Desc"));

        Assert.Equal("RECIBO_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Anular_ReciboEmitido_CambiaEstadoAnulado()
    {
        // BR-RECIBO-003
        var recibo = ReciboIngreso.Registrar("2600000001", DateTime.Today, 100m, "Test");
        recibo.Anular();

        Assert.Equal("04", recibo.Estado);
        Assert.Equal("ANULADO", recibo.DescripcionEstado);
        Assert.True(recibo.EstaAnulado);
    }

    [Fact]
    public void Anular_ReciboYaAnulado_LanzaExcepcion()
    {
        // BR-RECIBO-003
        var recibo = ReciboIngreso.Registrar("2600000001", DateTime.Today, 100m, "Test");
        recibo.Anular();

        var ex = Assert.Throws<DomainException>(() => recibo.Anular());
        Assert.Equal("RECIBO_YA_ANULADO", ex.Code);
    }

    [Fact]
    public void Anular_ReciboPagado_LanzaExcepcion()
    {
        // BR-RECIBO-003: no se puede anular un recibo pagado
        var recibo = ReciboIngreso.Reconstruir(
            "2600000001", DateTime.Today, 100m, "Test",
            "01", "01", null, null, 1m, null, false,
            "02", null, null, null, DateTime.Now, null, null);

        var ex = Assert.Throws<DomainException>(() => recibo.Anular());
        Assert.Equal("RECIBO_PAGADO_NO_ANULABLE", ex.Code);
    }

    [Fact]
    public void Registrar_ConAnticipo_MarcaFlag()
    {
        // BR-RECIBO-005
        var recibo = ReciboIngreso.Registrar(
            "2600000001", DateTime.Today, 200m, "Anticipo pedido",
            esAnticipo: true);

        Assert.True(recibo.EsAnticipo);
    }

    [Fact]
    public void DescripcionEstado_Pagado_RetornaTexto()
    {
        // BR-RECIBO-002
        var recibo = ReciboIngreso.Reconstruir(
            "2600000001", DateTime.Today, 100m, "Test",
            "01", "01", null, null, 1m, null, false,
            "02", null, null, null, DateTime.Now, null, null);

        Assert.Equal("PAGADO", recibo.DescripcionEstado);
        Assert.True(recibo.EstaPagado);
    }
}
