using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Maestros;

/// <summary>
/// Tests de dominio para CuentaCorriente.
/// Legacy: TDELIVERY (lClienteCtaCte=1) / vCompania.
/// BR-CTACTE-001, BR-CTACTE-002, BR-CTACTE-003.
/// </summary>
public class CuentaCorrienteTests
{
    [Fact]
    public void Crear_CuentaValida_CreaEntidad()
    {
        var ctacte = CuentaCorriente.Crear("DLV001", "Empresa Test", 1000m, 300m, "user");

        Assert.Equal("DLV001", ctacte.CodigoCtaCte);
        Assert.Equal("Empresa Test", ctacte.Descripcion);
        Assert.Equal(1000m, ctacte.LineaCredito);
        Assert.Equal(300m, ctacte.Consumo);
        Assert.Equal(700m, ctacte.Saldo);
    }

    [Fact]
    public void Saldo_CalculadoCorrectamente()
    {
        var ctacte = CuentaCorriente.Crear("DLV001", "Empresa Test", 500m, 200m, "user");
        Assert.Equal(300m, ctacte.Saldo);
    }

    [Fact]
    public void AplicarConsumo_DentroDeLinea_Actualiza()
    {
        var ctacte = CuentaCorriente.Crear("DLV001", "Empresa Test", 1000m, 100m, "user");
        ctacte.AplicarConsumo(200m);
        Assert.Equal(300m, ctacte.Consumo);
        Assert.Equal(700m, ctacte.Saldo);
    }

    [Fact]
    public void AplicarConsumo_SuperaLinea_LanzaDomainException()
    {
        var ctacte = CuentaCorriente.Crear("DLV001", "Empresa Test", 500m, 400m, "user");
        var ex = Assert.Throws<DomainException>(() => ctacte.AplicarConsumo(200m));
        Assert.Equal("CTACTE_SALDO_INSUFICIENTE", ex.Code);
    }

    [Fact]
    public void Crear_LineaCredito_Negativa_LanzaDomainException()
    {
        var ex = Assert.Throws<DomainException>(() =>
            CuentaCorriente.Crear("DLV001", "Empresa Test", -100m, 0m, "user"));
        Assert.Equal("CTACTE_LINEA_INVALIDA", ex.Code);
    }

    [Fact]
    public void Crear_CodigoVacio_LanzaDomainException()
    {
        var ex = Assert.Throws<DomainException>(() =>
            CuentaCorriente.Crear("", "Empresa Test", 100m, 0m, "user"));
        Assert.Equal("CTACTE_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void TieneSaldoDisponible_MontoValido_RetornaTrue()
    {
        var ctacte = CuentaCorriente.Crear("DLV001", "Empresa Test", 1000m, 300m, "user");
        Assert.True(ctacte.TieneSaldoDisponible(600m));
    }

    [Fact]
    public void TieneSaldoDisponible_MontoExcede_RetornaFalse()
    {
        var ctacte = CuentaCorriente.Crear("DLV001", "Empresa Test", 1000m, 900m, "user");
        Assert.False(ctacte.TieneSaldoDisponible(200m));
    }
}
