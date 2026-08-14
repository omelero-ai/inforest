using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Caja;

/// <summary>
/// Tests de dominio para ReciboEgreso.
/// Cubre BR-RECIBO-007, BR-RECIBO-008, BR-RECIBO-010, BR-RECIBO-011, BR-RECIBO-012.
/// </summary>
public class ReciboEgresoTests
{
    [Fact]
    public void Registrar_ConDatosValidos_RetornaEgresoEmitido()
    {
        var recibo = ReciboEgreso.Registrar(
            "2600000001", DateTime.Today, 80m, "Pago proveedor", "VARIOS", tipoEgreso: "GEN");

        Assert.Equal("2600000001", recibo.CodigoRecibo);
        Assert.Equal(80m, recibo.Monto);
        Assert.Equal("01", recibo.Estado);
        Assert.Equal("EMITIDO", recibo.DescripcionEstado);
        Assert.Equal("VARIOS", recibo.ModoEgreso);
        Assert.True(recibo.EstaEmitido);
    }

    [Fact]
    public void Registrar_MontoNegativo_LanzaExcepcion()
    {
        // BR-RECIBO-011
        var ex = Assert.Throws<DomainException>(() =>
            ReciboEgreso.Registrar("2600000001", DateTime.Today, -10m, "Desc"));

        Assert.Equal("EGRESO_MONTO_INVALIDO", ex.Code);
    }

    [Fact]
    public void Registrar_DescripcionVacia_LanzaExcepcion()
    {
        // BR-RECIBO-011
        var ex = Assert.Throws<DomainException>(() =>
            ReciboEgreso.Registrar("2600000001", DateTime.Today, 50m, ""));

        Assert.Equal("EGRESO_DESCRIPCION_REQUERIDA", ex.Code);
    }

    [Fact]
    public void Registrar_ModoNotaCreditoSinReferencia_LanzaExcepcion()
    {
        // BR-RECIBO-012
        var ex = Assert.Throws<DomainException>(() =>
            ReciboEgreso.Registrar(
                "2600000001", DateTime.Today, 50m, "NC egreso", "NOTACREDITO", referencia: null));

        Assert.Equal("EGRESO_REFERENCIA_NC_REQUERIDA", ex.Code);
    }

    [Fact]
    public void Registrar_ModoNotaCreditoConReferencia_Exitoso()
    {
        // BR-RECIBO-012: con referencia ok
        var recibo = ReciboEgreso.Registrar(
            "2600000001", DateTime.Today, 50m, "NC egreso", "NOTACREDITO",
            referencia: "NC2600000001");

        Assert.Equal("NOTACREDITO", recibo.ModoEgreso);
        Assert.Equal("NC2600000001", recibo.Referencia);
    }

    [Fact]
    public void Registrar_ModoInvalido_LanzaExcepcion()
    {
        // BR-RECIBO-010
        var ex = Assert.Throws<DomainException>(() =>
            ReciboEgreso.Registrar("2600000001", DateTime.Today, 50m, "Desc", "INVALIDO"));

        Assert.Equal("EGRESO_MODO_INVALIDO", ex.Code);
    }

    [Fact]
    public void Anular_ReciboEmitido_CambiaEstadoAnulado()
    {
        // BR-RECIBO-008
        var recibo = ReciboEgreso.Registrar("2600000001", DateTime.Today, 80m, "Test");
        recibo.Anular();

        Assert.Equal("04", recibo.Estado);
        Assert.Equal("ANULADO", recibo.DescripcionEstado);
        Assert.True(recibo.EstaAnulado);
    }

    [Fact]
    public void Anular_ReciboYaAnulado_LanzaExcepcion()
    {
        // BR-RECIBO-008
        var recibo = ReciboEgreso.Registrar("2600000001", DateTime.Today, 80m, "Test");
        recibo.Anular();

        var ex = Assert.Throws<DomainException>(() => recibo.Anular());
        Assert.Equal("EGRESO_YA_ANULADO", ex.Code);
    }
}
