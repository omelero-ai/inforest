using Inforest.Domain.Entities.Hardware;

namespace Inforest.Infrastructure.Tests.Hardware;

/// <summary>
/// Pruebas de serialización/validación de DocumentoFE y entidades de dominio de periféricos.
/// Regla BR-PERIPH-006.
/// </summary>
public sealed class DomainHardwareEntityTests
{
    [Fact]
    public void DocumentoFE_Crear_EstadoInicialEsPendiente()
    {
        var doc = DocumentoFE.Crear(
            "DOC-001", "FACTURA", "F001", "00000001", "PE",
            "20123456789", "Empresa SA",
            "12345678", "Cliente",
            [new ItemDocumentoFE("P01", "Pizza", 1, 50.00m, 50.00m, 9.00m)],
            50.00m, 9.00m, 59.00m, DateTime.Now);

        Assert.Equal(EstadoDocumentoFE.Pendiente, doc.Estado);
        Assert.Null(doc.NumeroAutorizacion);
    }

    [Fact]
    public void DocumentoFE_RegistrarAutorizacion_EstadoEsAutorizado()
    {
        var doc = DocumentoFE.Crear(
            "DOC-001", "FACTURA", "F001", "00000001", "PE",
            "20123456789", "Empresa SA",
            "12345678", "Cliente",
            [new ItemDocumentoFE("P01", "Pizza", 1, 50.00m, 50.00m, 9.00m)],
            50.00m, 9.00m, 59.00m, DateTime.Now);

        doc.RegistrarAutorizacion("12345678901234567890", "QR-DATA", "<xml/>");

        Assert.Equal(EstadoDocumentoFE.Autorizado, doc.Estado);
        Assert.Equal("12345678901234567890", doc.NumeroAutorizacion);
        Assert.Equal("QR-DATA", doc.QrData);
    }

    [Fact]
    public void DocumentoFE_SinItems_LanzaArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            DocumentoFE.Crear(
                "DOC-001", "FACTURA", "F001", "00000001", "PE",
                "20123456789", "Empresa SA",
                "12345678", "Cliente",
                [],
                0.00m, 0.00m, 0.00m, DateTime.Now));
    }

    [Fact]
    public void BlueVisionTicket_Crear_SetMozoSinMozoSiNulo()
    {
        var ticket = BlueVisionTicket.Crear("T01", "Salon", "Mesa 1", null!, "EVENTO", 100.00m, []);
        Assert.Equal("SIN MOZO", ticket.Mozo);
    }

    [Fact]
    public void BlueVisionLinea_Crear_TotalCalculado()
    {
        var linea = BlueVisionLinea.Crear("L01", "T01", "Pizza", 25.00m, 2);
        Assert.Equal(50.00m, linea.Total);
    }

    [Fact]
    public void ConfiguracionPais_PorCodigo_CodigoDesconocido_RetornaPeru()
    {
        var config = ConfiguracionPais.PorCodigo("ZZ");
        Assert.Equal("PE", config.CodigoPais);
    }

    [Fact]
    public void ConfiguracionPais_Argentina_RequiereFiscalTrue()
    {
        Assert.True(ConfiguracionPais.Argentina.RequiereFiscal);
    }

    [Fact]
    public void FiscalDocumento_Crear_SinItems_LanzaArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            FiscalDocumento.Crear(
                "A", "Empresa SA", "Empresa SA",
                "30-12345678-9", "Av. Corrientes 123", "4321-5678",
                "0001-00000001", new DateOnly(2026, 1, 12),
                null, null, [], 0));
    }
}
