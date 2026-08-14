using Inforest.Application.Configuracion;
using Inforest.Application.Turno;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Entities.Configuracion;
using Moq;

namespace Inforest.Application.Tests.Turno;

/// <summary>
/// Tests para CerrarTurnoHandler.
/// Legacy: frmLiquidacionDetalle.frm — cmdOpcion_Click(0).
/// Reglas: BR-CAJA-001 (lObligaCierre), BR-CAJA-002 (lActivaConsultaDescargo), BR-CAJA-004 (UPDATE MTURNO).
/// </summary>
public class CerrarTurnoHandlerTests
{
    // ─────────────────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────────────────

    private static ITurnoRepository TurnoRepoOk()
    {
        var mock = new Mock<ITurnoRepository>();
        mock.Setup(r => r.CerrarAsync(
                It.IsAny<string>(), It.IsAny<decimal>(), It.IsAny<CierreTurnoBreakdown>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);
        return mock.Object;
    }

    private static IParametroRepository ConfigRepoNullDefaults()
    {
        var mock = new Mock<IParametroRepository>();
        mock.Setup(r => r.ObtenerConfiguracionCajaAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((ConfiguracionCaja?)null);
        mock.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync((ConfiguracionSistema?)null);
        return mock.Object;
    }

    private static ConfiguracionCaja BuildCaja(bool lObligaCierre)
        => ConfiguracionCaja.Cargar(
            "CA1", tDescripcion: null, tPrecuenta: null,
            lComanda: false, vComanda: false, lMotivoEliminaC: false, lMotivoElimina: false,
            lActivo: true, lRefresca: false, lPasswordC: false, lPassword: false,
            tGrupo: null, lConsumo1: false, lConsumo2: false, lConsumo3: false,
            lPrecuenta: false, lAdicion: false, lPrecuentaAgrupada: false,
            tTipoPedido: null, lObliga: false, lMozo: false, lObligaPrinter: false,
            lPax: false, lObligaCierre: lObligaCierre, lFiltroTipoPedido: false,
            nPuerto: null, tMensaje1: null, tMensaje2: null,
            lCancelacion: false, lDirecto: false, lObligaPrecuenta: false,
            lComboPrecuenta: false, lComboDocumento: false, lCambioMesa: false,
            lVisaNet: false, lImpuestoPrecuenta: false, lDocumentoAgrupado: false,
            lOrden: false, lValorCortesia: false, lObservacion: false,
            lCajaRapida: false, lPropiedadPrecuenta: false, lPropiedadDocumento: false,
            lPrecioNetoPrecuenta: false, nLimitePrecuenta: null, tUnidadNegocio: null,
            nLimiteReimpresion: null, lPasswordTransferencia: false, lCD: false,
            lFechaEntregaDelivery: false, lMultiCajero: false, lMCPV: false,
            lCCVOX: false, lMotorizado: false, lEquivaDolaPrecuenta: false,
            tSubAlmacen: null, lObservacionPrecuenta: false, lObservacionDocumento: false,
            lPasswordImportarPedido: false, lActivaImpDscAlternativa: false,
            lCompatibilidadTVS: false, nLongitudBarra: null, lPagoRapido: false,
            lDisgrega: false, lPasswordPorCobrar: false, lModificaTipoPedido: false,
            tSucursal: null, nBalanzaPuerto: null, lCapturaPeso: false,
            lPagoRapidoPV: false, tTextoConsumo: null,
            
            lSiab: false, tSectorVenta: null, lCajaMobile: false,
            lBloqueaPrecuenta: false, lRotulado: false,
            lMultiAreaSubGrupo: false, lMultiAreaCaja: false,
            lHuella: false, lImprimeImagCabPrecuenta: false, lImprimeImagPiePrecuenta: false,
            lAccesoDespachoPedido: false, lBuscaPedidoNumero: false,
            lCodigoReciboIngreso: false, lPagoRapidoMod: false,
            lConsumo4: false, lPrecuentaNoValorizada: false, lWebAp: false,
            lBuscarPedidoVisualizarGrilla: false, lBuscarPedidoFiltrarMesa: false,
            lClaveEnvioProduccion: false, lMesa247: false, EAN13: false,
            lObservacionCabDoc: false, tCompMesa247: null, lMotivoDescuento: false,
            lCajaContingencia: false, lImpPropina: false, lImpcomandaf2: false,
            lPassOtrosPagos: false, lImpClienteCab: false, tImpFiscal: null);

    private static ConfiguracionSistema BuildSistema(bool lActivaConsultaDescargo)
        => ConfiguracionSistema.Cargar(
            tIdentificacionTributaria: null, tRazonSocial: null, tRazonComercial: null,
            tDireccion: null, tTelefono: null, tEmail: null, tWebPage: null,
            tMonedaN: null, tMonN: null, tMonedaE: null, tMonE: null,
            nTiempo: null, nChkTiempo: null, Impuesto1: 0d, Impuesto2: 0d, Impuesto3: 0d,
            tImpuesto1: null, tImpuesto2: null, tImpuesto3: null, nCorrelativo: null,
            nDelivery: 0d, nLlevar: 0d, nCanal4: 0d, nCanal5: 0d, tPie: null,
            lBotonTrans: false, tElimina: null, tPassword: null, nItem: null,
            lLongitud: false, nLongitud: null, lPrinter: false, lAlmacen: false, lRapido: false,
            tBoton1: null, tBoton2: null, tBoton3: null, tBoton4: null, tBoton5: null,
            tPiePreCuenta: null, lInfhotel: false, tClub: null, nPunto: 0d, lCierre: false,
            nDecimal: null, nDias: null, lEquivalencia: false, nCabecera: null, nDetalle: null,
            tVersion: null, lComboGeneral: false, nDiasDelivery: null, nTiempoMinutoCD: null,
            lMultilocal: false, lKDS: false, tOrderInfo: null, tOrderStatus: null, tBump: null,
            lDiaContableAutomatico: false, tHoraCierreDiaContable: null, lDiaContableManual: false,
            lClub: false, lImprimeDiaContable: false, nItemGuia: null, nCabeceraGuia: null,
            nDetalleGuia: null, nAsignacionMotorizado: 0d, tTarifaActualMotorizado: null,
            tUsuarioTarifa: null, fRegistroTarifa: null, lEnvioChef: false,
            tContribuyenteEspecial: null, fContribuyenteEspecial: null,
            tDireccion2: null, lMobileUnidadNegocio: false, lMobilePasswordCCaja: false,
            lActivaConsultaDescargo: lActivaConsultaDescargo,
            nCabeceraV: null, nItemV: null, nPieV: null,
            lFacturacionE: false, lControlUsuario: false, lHoraEntregaDelivery: false,
            tCodigoFE: null, tPieDocumento1: null, lAmbienteFE: false,
            lHuellaDigital: false, lHuellaSecugen: false, tAgenteRetencion: null, tFax: null,
            lLogCajaRapida: false, nTiempoToleranciaAnf: null,
            lEmailConfirmacion: false, tEmailConfirmacion: null,
            lEmailRecordatorio: false, tEmailRecordatorio: null,
            lEmailAgradecimiento: false, tEmailAgradecimiento: null,
            lImprimeCodigoBarras: false, lEnvioAutomatico: false,
            lControlEnviosProduccion: false, lActivaTarjeta: false,
            lEventos: false, lFEOfisis: false, tCodigoEmpresa: null, tCodigoTienda: null,
            tCodigoMarca: null, tCodigoUbigeo: null, lPagoAntesImpresion: false,
            lCambioContrasena: false, tCajaMesa247: null, tAdicionMesa247: null,
            tServidorFE: null, tBDFE: null, lActivaCuenCorrienteAut: false,
            lInteSAP: false, tServidorSAP: null, tBDSAP: null, tCodAlmcSAP: null,
            tCuentaContableCort: null, lVisor8: false, lvisortactil: false,
            lvisortiempo: null, lGlosaTransGratuita: null, lActivaCover: false,
            tMontoMinCover: null, tCodItemCover: null, lNcOfisis: false, tRutaFE: null,
            lCodigoQrFE: false, lactivaFechaNC: false, lCheffFiltroSalon: false,
            lCheffFiltroSubGrupo: false, lFEpape: false, lDesPagoCheque: false,
            lDesPagoOtro: false, tUsuarioFE: null, tClaveFE: null,
            lFESpring: false, lFECarbajal: false, lAnula: false, paramCarvajal: null,
            tCarvajalCorreos: null, lParcialNC: false, lNCElimina: false,
            lValidaDNI: false, lVerTCImp: false, lFETCI: false, lDesactivaNCFP: false,
            lFEBiz: false, tCodAnticipo: null, lActivaAnticipo: false,
            lFEGood: false, tMaxMotorizado: null, lStockDescargo: false,
            lFEubl21: false, lBloqInafecto: false, lDescargoInsumo: false,
            lEstupendoFE: false, lFEGesa: false,
            tMonedaE2: null, tMonE2: null, tMonedaE3: null, tMonE3: null);

    // ─────────────────────────────────────────────────────────────────────────
    // Cierre normal
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task CerrarTurno_Normal_Exitoso()
    {
        var handler = new CerrarTurnoHandler(TurnoRepoOk(), ConfigRepoNullDefaults());
        var result = await handler.HandleAsync(new CerrarTurnoCommand("TUR001", "CA1", 500m));
        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task CerrarTurno_CodigoTurnoVacio_Falla()
    {
        var handler = new CerrarTurnoHandler(TurnoRepoOk(), ConfigRepoNullDefaults());
        var result = await handler.HandleAsync(new CerrarTurnoCommand("", "CA1", 0m));
        Assert.False(result.EsExitoso);
        Assert.Equal("TURNO_CODIGO_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task CerrarTurno_CodigoCajaVacio_Falla()
    {
        var handler = new CerrarTurnoHandler(TurnoRepoOk(), ConfigRepoNullDefaults());
        var result = await handler.HandleAsync(new CerrarTurnoCommand("TUR001", "", 0m));
        Assert.False(result.EsExitoso);
        Assert.Equal("TURNO_CAJA_REQUERIDA", result.CodigoError);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // BR-CAJA-001: lObligaCierre
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task CerrarTurno_ObligaCierre_SinSupervisor_RequiereSupervisor()
    {
        var configRepo = new Mock<IParametroRepository>();
        configRepo.Setup(r => r.ObtenerConfiguracionCajaAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildCaja(lObligaCierre: true));
        configRepo.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync((ConfiguracionSistema?)null);

        var handler = new CerrarTurnoHandler(TurnoRepoOk(), configRepo.Object);
        var result = await handler.HandleAsync(new CerrarTurnoCommand("TUR001", "CA1", 500m));

        Assert.False(result.EsExitoso);
        Assert.Equal("REQUIERE_SUPERVISOR", result.CodigoError);
    }

    [Fact]
    public async Task CerrarTurno_ObligaCierre_ConSupervisor_Exitoso()
    {
        var configRepo = new Mock<IParametroRepository>();
        configRepo.Setup(r => r.ObtenerConfiguracionCajaAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildCaja(lObligaCierre: true));
        configRepo.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync((ConfiguracionSistema?)null);

        var handler = new CerrarTurnoHandler(TurnoRepoOk(), configRepo.Object);
        var result = await handler.HandleAsync(
            new CerrarTurnoCommand("TUR001", "CA1", 500m, SupervisorAutorizado: true));

        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task CerrarTurno_ObligaCierre_False_SinSupervisor_Exitoso()
    {
        var handler = new CerrarTurnoHandler(TurnoRepoOk(), ConfigRepoNullDefaults());
        var result = await handler.HandleAsync(new CerrarTurnoCommand("TUR001", "CA1", 500m));
        Assert.True(result.EsExitoso);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // BR-CAJA-002: lActivaConsultaDescargo
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task CerrarTurno_ActivaDescargo_SinConfirmar_RequiereConfirmacion()
    {
        var configRepo = new Mock<IParametroRepository>();
        configRepo.Setup(r => r.ObtenerConfiguracionCajaAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((ConfiguracionCaja?)null);
        configRepo.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildSistema(lActivaConsultaDescargo: true));

        var handler = new CerrarTurnoHandler(TurnoRepoOk(), configRepo.Object);
        var result = await handler.HandleAsync(new CerrarTurnoCommand("TUR001", "CA1", 500m));

        Assert.False(result.EsExitoso);
        Assert.Equal("REQUIERE_CONFIRMACION_DESCARGO", result.CodigoError);
    }

    [Fact]
    public async Task CerrarTurno_ActivaDescargo_Confirmado_Exitoso()
    {
        var configRepo = new Mock<IParametroRepository>();
        configRepo.Setup(r => r.ObtenerConfiguracionCajaAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((ConfiguracionCaja?)null);
        configRepo.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(BuildSistema(lActivaConsultaDescargo: true));

        var handler = new CerrarTurnoHandler(TurnoRepoOk(), configRepo.Object);
        var result = await handler.HandleAsync(
            new CerrarTurnoCommand("TUR001", "CA1", 500m, DescargoPendienteConfirmado: true));

        Assert.True(result.EsExitoso);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // BR-CAJA-004: Desglose de montos
    // ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task CerrarTurno_ConBreakdown_LlamaCerrarConDesglose()
    {
        var turnoRepo = new Mock<ITurnoRepository>();
        CierreTurnoBreakdown? captured = null;
        turnoRepo.Setup(r => r.CerrarAsync(
                "TUR001", It.IsAny<decimal>(), It.IsAny<CierreTurnoBreakdown>(), It.IsAny<CancellationToken>()))
            .Callback<string, decimal, CierreTurnoBreakdown, CancellationToken>((_, _, b, _) => captured = b)
            .ReturnsAsync(true);

        var breakdown = new CierreTurnoBreakdown
        {
            EfectivoMN = 200m, EfectivoME = 50m, ChequesMN = 100m, RetiroMN = 20m
        };

        var handler = new CerrarTurnoHandler(turnoRepo.Object, ConfigRepoNullDefaults());
        var result = await handler.HandleAsync(
            new CerrarTurnoCommand("TUR001", "CA1", 180m, Breakdown: breakdown));

        Assert.True(result.EsExitoso);
        Assert.NotNull(captured);
        Assert.Equal(200m, captured!.EfectivoMN);
        Assert.Equal(50m, captured.EfectivoME);
        Assert.Equal(100m, captured.ChequesMN);
        Assert.Equal(180m, captured.MontoFinalMN); // 200 - 20
    }

    [Fact]
    public async Task CerrarTurno_SinBreakdown_UsaBreakdownVacio()
    {
        CierreTurnoBreakdown? captured = null;
        var turnoRepo = new Mock<ITurnoRepository>();
        turnoRepo.Setup(r => r.CerrarAsync(
                It.IsAny<string>(), It.IsAny<decimal>(), It.IsAny<CierreTurnoBreakdown>(), It.IsAny<CancellationToken>()))
            .Callback<string, decimal, CierreTurnoBreakdown, CancellationToken>((_, _, b, _) => captured = b)
            .ReturnsAsync(true);

        var handler = new CerrarTurnoHandler(turnoRepo.Object, ConfigRepoNullDefaults());
        await handler.HandleAsync(new CerrarTurnoCommand("TUR001", "CA1", 0m));

        Assert.NotNull(captured);
        Assert.Equal(0m, captured!.EfectivoMN);
        Assert.Equal(0m, captured.MontoFinalMN);
        Assert.Empty(captured.Tarjetas);
    }
}
