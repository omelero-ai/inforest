using Inforest.Application.Interfaces.Hardware;
using Inforest.Infrastructure.Country;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace Inforest.Infrastructure.Tests.Hardware;

/// <summary>
/// Pruebas unitarias de FacturacionElectronicaFactory y gateways Null/país.
/// Regla BR-PERIPH-006, BR-PERIPH-007. ADR-010.
/// </summary>
public sealed class FacturacionElectronicaFactoryTests
{
    [Fact]
    public void ObtenerGateway_CodigoPaisConocido_RetornaGatewayCorresto()
    {
        var gateways = new IFacturacionElectronicaGateway[]
        {
            new NullFEGateway(NullLogger<NullFEGateway>.Instance),
            new PeruFEGateway(NullLogger<PeruFEGateway>.Instance, Options.Create(new PeruFEOptions())),
            new ArgentinaFEGateway(NullLogger<ArgentinaFEGateway>.Instance, Options.Create(new ArgentinaFEOptions()))
        };
        var factory = new FacturacionElectronicaFactory(gateways, NullLogger<FacturacionElectronicaFactory>.Instance);

        var gatewayPE = factory.ObtenerGateway("PE");
        var gatewayAR = factory.ObtenerGateway("AR");

        Assert.Equal("PE", gatewayPE.CodigoPais);
        Assert.Equal("AR", gatewayAR.CodigoPais);
    }

    [Fact]
    public void ObtenerGateway_CodigoPaisDesconocido_RetornaNullGateway()
    {
        var gateways = new IFacturacionElectronicaGateway[]
        {
            new NullFEGateway(NullLogger<NullFEGateway>.Instance)
        };
        var factory = new FacturacionElectronicaFactory(gateways, NullLogger<FacturacionElectronicaFactory>.Instance);

        var gateway = factory.ObtenerGateway("XX");

        Assert.Equal("NULL", gateway.CodigoPais);
    }

    [Fact]
    public async Task NullFEGateway_EmitirDocumento_EsExitoso()
    {
        var gateway = new NullFEGateway(NullLogger<NullFEGateway>.Instance);
        var doc = CrearDocumentoFEDto("PE");

        var result = await gateway.EmitirDocumentoAsync(doc);

        Assert.True(result.EsExitoso);
        Assert.NotNull(result.NumeroAutorizacion);
    }

    [Fact]
    public async Task NullFEGateway_AnularDocumento_EsExitoso()
    {
        var gateway = new NullFEGateway(NullLogger<NullFEGateway>.Instance);
        var result = await gateway.AnularDocumentoAsync("DOC-001", "Anulación de prueba");
        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task NullFEGateway_ConsultarEstado_EsExitoso()
    {
        var gateway = new NullFEGateway(NullLogger<NullFEGateway>.Instance);
        var result = await gateway.ConsultarEstadoAsync("DOC-001");
        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task PeruFEGateway_EmitirDocumento_CodigoPaisEsPE()
    {
        var gateway = new PeruFEGateway(
            NullLogger<PeruFEGateway>.Instance,
            Options.Create(new PeruFEOptions()));
        Assert.Equal("PE", gateway.CodigoPais);

        var doc = CrearDocumentoFEDto("PE");
        var result = await gateway.EmitirDocumentoAsync(doc);
        // Resultado simulado esperado (implementación real pendiente)
        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task ArgentinaFEGateway_AnularDocumento_RetornaFallido_PorqueAFIPNoAnula()
    {
        var gateway = new ArgentinaFEGateway(
            NullLogger<ArgentinaFEGateway>.Instance,
            Options.Create(new ArgentinaFEOptions()));

        var result = await gateway.AnularDocumentoAsync("CAE-001", "Error de carga");

        Assert.False(result.EsExitoso);
        Assert.Equal("AFIP_NO_ANULACION", result.CodigoRespuesta);
    }

    private static DocumentoFEDto CrearDocumentoFEDto(string codigoPais) =>
        new DocumentoFEDto(
            "FACTURA",
            "F001",
            "00000001",
            codigoPais,
            new EmisorFEDto("20123456789", "Empresa Test SA", "Av. Lima 123"),
            new ReceptorFEDto("12345678", "Cliente Test"),
            [new ItemFEDto("P001", "Pizza Margarita", 1, 50.00m, 50.00m, 9.00m)],
            50.00m,
            9.00m,
            59.00m);
}
