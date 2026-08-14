using Inforest.Application.Interfaces.Country;
using Inforest.Infrastructure.Country;

namespace Inforest.Infrastructure.Tests.Hardware;

/// <summary>
/// Pruebas de PaisPolicyFactory y políticas de país.
/// Regla BR-PERIPH-007. ADR-008.
/// </summary>
public sealed class PaisPolicyTests
{
    private static PaisPolicyFactory CrearFactory() =>
        new PaisPolicyFactory(new IPaisPolicy[]
        {
            new PeruPaisPolicy(),
            new ArgentinaPaisPolicy(),
            new ChilePaisPolicy(),
            new EcuadorPaisPolicy(),
            new BoliviaPaisPolicy(),
            new EspanaPaisPolicy()
        });

    [Theory]
    [InlineData("PE", "R.U.C.", "SUNAT", "PEN", false, true)]
    [InlineData("AR", "C.U.I.T.", "AFIP", "ARS", true, true)]
    [InlineData("CL", "R.U.T.", "SII", "CLP", false, true)]
    [InlineData("EC", "R.U.C.", "SRI", "USD", false, true)]
    [InlineData("BO", "N.I.T.", "SIAT", "BOB", false, true)]
    [InlineData("ES", "N.I.F.", "AEAT", "EUR", false, false)]
    public void ObtenerPolitica_PaisConocido_RetornaPropiedadesCorrectas(
        string codigoPais, string labelDoc, string tipoFE, string moneda,
        bool requiereFiscal, bool requiereFE)
    {
        var factory = CrearFactory();
        var politica = factory.ObtenerPolitica(codigoPais);

        Assert.Equal(codigoPais, politica.CodigoPais);
        Assert.Equal(labelDoc, politica.LabelDocumentoTributario);
        Assert.Equal(tipoFE, politica.TipoFEGateway);
        Assert.Equal(moneda, politica.CodigoMoneda);
        Assert.Equal(requiereFiscal, politica.RequiereFiscal);
        Assert.Equal(requiereFE, politica.RequiereFE);
    }

    [Fact]
    public void ObtenerPolitica_PaisDesconocido_RetornaPeruPorDefecto()
    {
        // Legacy: comportamiento por defecto = Perú cuando `pais` no está reconocido
        var factory = CrearFactory();
        var politica = factory.ObtenerPolitica("ZZ");
        Assert.Equal("PE", politica.CodigoPais);
    }

    [Fact]
    public void ObtenerPolitica_CodigoInsensibleMayusculas()
    {
        var factory = CrearFactory();
        var politica = factory.ObtenerPolitica("pe");
        Assert.Equal("PE", politica.CodigoPais);
    }

    [Fact]
    public void ArgentinaPaisPolicy_RequiereFiscalTrue()
    {
        // Legacy: Argentina es el único país que usa IFEpson.ocx (pais = "003")
        var politica = new ArgentinaPaisPolicy();
        Assert.True(politica.RequiereFiscal);
    }
}
