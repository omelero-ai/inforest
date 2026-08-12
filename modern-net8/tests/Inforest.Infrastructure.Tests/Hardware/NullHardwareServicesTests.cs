using Inforest.Infrastructure.Hardware;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace Inforest.Infrastructure.Tests.Hardware;

/// <summary>
/// Pruebas unitarias de los servicios Null (stubs) de hardware.
/// Verifica que los stubs no lancen excepción y retornen resultados válidos.
/// Reglas BR-PERIPH-001 a BR-PERIPH-005.
/// </summary>
public sealed class NullHardwareServicesTests
{
    // --- PinPad ---

    [Fact]
    public async Task NullPinPadService_InicializarAsync_NoLanzaExcepcion()
    {
        var sut = new NullPinPadService(NullLogger<NullPinPadService>.Instance);
        var result = await sut.InicializarAsync("test.ini");
        Assert.True(result.EsExitoso);
        Assert.Equal(0, result.CodigoRetorno);
    }

    [Fact]
    public async Task NullPinPadService_EjecutarVenta_RetornaOk()
    {
        var sut = new NullPinPadService(NullLogger<NullPinPadService>.Instance);
        var result = await sut.EjecutarVentaAsync(100.00m, 1);
        Assert.True(result.EsExitoso);
        Assert.NotNull(result.DatosAutorizacion);
    }

    [Fact]
    public async Task NullPinPadService_TestComunicacion_RetornaOk()
    {
        var sut = new NullPinPadService(NullLogger<NullPinPadService>.Instance);
        var result = await sut.TestComunicacionAsync();
        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task NullPinPadService_Anular_RetornaOk()
    {
        var sut = new NullPinPadService(NullLogger<NullPinPadService>.Instance);
        var result = await sut.AnularAsync("OP-001");
        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task NullPinPadService_CerrarPuerto_NoLanzaExcepcion()
    {
        var sut = new NullPinPadService(NullLogger<NullPinPadService>.Instance);
        await sut.CerrarPuertoAsync();
    }

    // --- CashDro ---

    [Fact]
    public async Task NullCashDroService_IniciarMotor_NoLanzaExcepcion()
    {
        var sut = new NullCashDroService(NullLogger<NullCashDroService>.Instance);
        await sut.IniciarMotorAsync("ABRIR");
    }

    [Fact]
    public async Task NullCashDroService_AbrirCajon_NoLanzaExcepcion()
    {
        var sut = new NullCashDroService(NullLogger<NullCashDroService>.Instance);
        await sut.AbrirCajonAsync();
    }

    // --- BlueVision ---

    [Fact]
    public async Task NullBlueVisionService_EnviarTicket_RetornaTrue()
    {
        var sut = new NullBlueVisionService(NullLogger<NullBlueVisionService>.Instance);
        var ticket = new Inforest.Application.Interfaces.Hardware.BlueVisionTicketDto(
            "guid-001", "Salon", "Mesa 1", "Mozo", "NUEVO_PEDIDO", 100.00m,
            new List<Inforest.Application.Interfaces.Hardware.BlueVisionLineaDto>());

        var result = await sut.EnviarTicketAsync(ticket);
        Assert.True(result);
    }

    // --- Biometría ---

    [Fact]
    public async Task NullBiometriaService_CapturarHuella_EsExitoso()
    {
        var sut = new NullBiometriaService(NullLogger<NullBiometriaService>.Instance);
        var result = await sut.CapturarHuellaAsync();
        Assert.True(result.EsExitoso);
        Assert.Equal(0, result.CodigoError);
    }

    [Fact]
    public async Task NullBiometriaService_VerificarUsuario_EsExitoso()
    {
        var sut = new NullBiometriaService(NullLogger<NullBiometriaService>.Instance);
        var result = await sut.VerificarUsuarioAsync("USR001", Array.Empty<byte>());
        Assert.True(result.EsExitoso);
    }

    // --- Impresora Fiscal ---

    [Fact]
    public async Task NullImpresoraFiscalService_Configurar_RetornaTrue()
    {
        var sut = new NullImpresoraFiscalService(NullLogger<NullImpresoraFiscalService>.Instance);
        var config = new Inforest.Application.Interfaces.Hardware.ConfiguracionImpresoraFiscal("COM1");
        var result = await sut.ConfigurarAsync(config);
        Assert.True(result);
    }

    [Fact]
    public async Task NullImpresoraFiscalService_EmitirFactura_RetornaTrue()
    {
        var sut = new NullImpresoraFiscalService(NullLogger<NullImpresoraFiscalService>.Instance);
        var doc = new Inforest.Application.Interfaces.Hardware.FiscalDocumentoDto(
            "A", "Empresa SA", "Empresa SA", "30-12345678-9", "Av. Corrientes 123",
            "0001-00000001", new DateOnly(2026, 1, 12), null, null,
            [new Inforest.Application.Interfaces.Hardware.FiscalItemDto(1, "Pizza", 100.00m, 21.0m)],
            121.00m);
        var result = await sut.EmitirFacturaAsync(doc);
        Assert.True(result);
    }

    [Fact]
    public async Task NullImpresoraFiscalService_CerrarJornada_RetornaTrue()
    {
        var sut = new NullImpresoraFiscalService(NullLogger<NullImpresoraFiscalService>.Instance);
        var result = await sut.CerrarJornadaAsync();
        Assert.True(result);
    }
}
