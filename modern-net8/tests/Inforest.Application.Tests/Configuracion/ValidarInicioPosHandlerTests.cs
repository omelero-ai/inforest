using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Configuracion;
using Moq;

namespace Inforest.Application.Tests.Configuracion;

public sealed class ValidarInicioPosHandlerTests
{
    [Fact]
    public async Task SinCodigoCaja_RetornaError()
    {
        var parametroService = new Mock<IParametroService>();
        var repository = new Mock<IParametroRepository>();
        var handler = new ValidarInicioPosHandler(parametroService.Object, repository.Object);

        var result = await handler.HandleAsync(new ValidarInicioPosQuery(string.Empty));

        Assert.False(result.EsExitoso);
        Assert.Equal("POS_CODIGO_CAJA_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task SinConfiguracionGlobal_RetornaError()
    {
        var parametroService = new Mock<IParametroService>();
        parametroService.Setup(x => x.TieneConfiguracionAsync(It.IsAny<CancellationToken>())).ReturnsAsync(false);
        var repository = new Mock<IParametroRepository>();
        var handler = new ValidarInicioPosHandler(parametroService.Object, repository.Object);

        var result = await handler.HandleAsync(new ValidarInicioPosQuery("001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("CONFIGURACION_SISTEMA_NO_ENCONTRADA", result.CodigoError);
    }

    [Fact]
    public async Task CajaNoExiste_RetornaError()
    {
        var parametroService = new Mock<IParametroService>();
        parametroService.Setup(x => x.TieneConfiguracionAsync(It.IsAny<CancellationToken>())).ReturnsAsync(true);
        var repository = new Mock<IParametroRepository>();
        repository.Setup(x => x.ObtenerConfiguracionCajaAsync("001", It.IsAny<CancellationToken>())).ReturnsAsync((ConfiguracionCaja?)null);
        var handler = new ValidarInicioPosHandler(parametroService.Object, repository.Object);

        var result = await handler.HandleAsync(new ValidarInicioPosQuery("001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("CONFIGURACION_CAJA_NO_ENCONTRADA", result.CodigoError);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task CajaValida_RespetaBypassMcpv(bool lMcpv, bool requiereLogin)
    {
        var parametroService = new Mock<IParametroService>();
        parametroService.Setup(x => x.TieneConfiguracionAsync(It.IsAny<CancellationToken>())).ReturnsAsync(true);
        var repository = new Mock<IParametroRepository>();
        repository
            .Setup(x => x.ObtenerConfiguracionCajaAsync("001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(CrearConfiguracionCaja("001", lMcpv));

        var handler = new ValidarInicioPosHandler(parametroService.Object, repository.Object);

        var result = await handler.HandleAsync(new ValidarInicioPosQuery("001"));

        Assert.True(result.EsExitoso);
        Assert.NotNull(result.Valor);
        Assert.Equal(requiereLogin, result.Valor!.RequiereLogin);
        Assert.Equal("001", result.Valor.CodigoCaja);
    }

    private static ConfiguracionCaja CrearConfiguracionCaja(string codigoCaja, bool lMcpv)
        => ConfiguracionCaja.Cargar(
            codigoCaja,
            tDescripcion: "Caja principal",
            tPrecuenta: null,
            lComanda: false,
            vComanda: false,
            lMotivoEliminaC: false,
            lMotivoElimina: false,
            lActivo: true,
            lRefresca: false,
            lPasswordC: false,
            lPassword: false,
            tGrupo: null,
            lConsumo1: false,
            lConsumo2: false,
            lConsumo3: false,
            lPrecuenta: false,
            lAdicion: false,
            lPrecuentaAgrupada: false,
            tTipoPedido: null,
            lObliga: false,
            lMozo: false,
            lObligaPrinter: false,
            lPax: false,
            lObligaCierre: false,
            lFiltroTipoPedido: false,
            nPuerto: null,
            tMensaje1: null,
            tMensaje2: null,
            lCancelacion: false,
            lDirecto: false,
            lObligaPrecuenta: false,
            lComboPrecuenta: false,
            lComboDocumento: false,
            lCambioMesa: false,
            lVisaNet: false,
            lImpuestoPrecuenta: false,
            lDocumentoAgrupado: false,
            lOrden: false,
            lValorCortesia: false,
            lObservacion: false,
            lCajaRapida: false,
            lPropiedadPrecuenta: false,
            lPropiedadDocumento: false,
            lPrecioNetoPrecuenta: false,
            nLimitePrecuenta: null,
            tUnidadNegocio: null,
            nLimiteReimpresion: null,
            lPasswordTransferencia: false,
            lCD: false,
            lFechaEntregaDelivery: false,
            lMultiCajero: false,
            lMCPV: lMcpv,
            lCCVOX: false,
            lMotorizado: false,
            lEquivaDolaPrecuenta: false,
            tSubAlmacen: null,
            lObservacionPrecuenta: false,
            lObservacionDocumento: false,
            lPasswordImportarPedido: false,
            lActivaImpDscAlternativa: false,
            lCompatibilidadTVS: false,
            nLongitudBarra: null,
            lPagoRapido: false,
            lDisgrega: false,
            lPasswordPorCobrar: false,
            lModificaTipoPedido: false,
            tSucursal: null,
            nBalanzaPuerto: null,
            lCapturaPeso: false,
            lPagoRapidoPV: false,
            tTextoConsumo: null,
            lSiab: false,
            tSectorVenta: null,
            lCajaMobile: false,
            lBloqueaPrecuenta: false,
            lRotulado: false,
            lMultiAreaSubGrupo: false,
            lMultiAreaCaja: false,
            lHuella: false,
            lImprimeImagCabPrecuenta: false,
            lImprimeImagPiePrecuenta: false,
            lAccesoDespachoPedido: false,
            lBuscaPedidoNumero: false,
            lCodigoReciboIngreso: false,
            lPagoRapidoMod: false,
            lConsumo4: false,
            lPrecuentaNoValorizada: false,
            lWebAp: false,
            lBuscarPedidoVisualizarGrilla: false,
            lBuscarPedidoFiltrarMesa: false,
            lClaveEnvioProduccion: false,
            lMesa247: false,
            EAN13: false,
            lObservacionCabDoc: false,
            tCompMesa247: null,
            lMotivoDescuento: false,
            lCajaContingencia: false,
            lImpPropina: false,
            lImpcomandaf2: false,
            lPassOtrosPagos: false,
            lImpClienteCab: false,
            tImpFiscal: null);
}
