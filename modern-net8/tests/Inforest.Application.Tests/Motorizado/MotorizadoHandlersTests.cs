using Inforest.Application.Motorizado;
using Inforest.Application.Configuracion;
using Inforest.Domain.Entities.Configuracion;
using Inforest.Domain.Entities.Motorizado;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Motorizado;

/// <summary>
/// Tests de handlers de Motorizado.
/// Legacy: frmAsignacionMotorizado.frm, frmLlegadaSalida.frm, frmTarifaMotorizado.frm
/// Reglas: BR-DEL-001, BR-DEL-002, BR-DEL-005, BR-DEL-006.
/// </summary>
public class MotorizadoHandlersTests
{
    [Fact]
    public async Task AsignarMotorizado_Todo_OK_RetornaOk()
    {
        var motorizado = Inforest.Domain.Entities.Motorizado.Motorizado.Crear("MOT001", "12345678", 10m, 15m, 20m);
        var pedido = Inforest.Domain.Entities.Delivery.PedidoDelivery.Crear(
            "PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);

        var mockMotRepo = new Mock<IMotorizadoRepository>();
        mockMotRepo.Setup(r => r.ObtenerPorCodigoAsync("MOT001", It.IsAny<CancellationToken>()))
                   .ReturnsAsync(motorizado);
        mockMotRepo.Setup(r => r.RegistrarAsignacionAsync(It.IsAny<AsignacionMotorizado>(), It.IsAny<CancellationToken>()))
                   .Returns(Task.CompletedTask);

        var mockPedRepo = new Mock<IPedidoDeliveryRepository>();
        mockPedRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
                   .ReturnsAsync(pedido);
        mockPedRepo.Setup(r => r.AsignarMotorizadoAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
                   .Returns(Task.CompletedTask);

        var handler = new AsignarMotorizadoHandler(mockMotRepo.Object, mockPedRepo.Object);
        var result = await handler.HandleAsync(new AsignarMotorizadoCommand("PED001", "MOT001", "USR001"));

        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task AsignarMotorizado_MotorizadoNoExiste_RetornaFallo()
    {
        var mockMotRepo = new Mock<IMotorizadoRepository>();
        mockMotRepo.Setup(r => r.ObtenerPorCodigoAsync("MOTXXX", It.IsAny<CancellationToken>()))
                   .ReturnsAsync((Inforest.Domain.Entities.Motorizado.Motorizado?)null);

        var mockPedRepo = new Mock<IPedidoDeliveryRepository>();

        var handler = new AsignarMotorizadoHandler(mockMotRepo.Object, mockPedRepo.Object);
        var result = await handler.HandleAsync(new AsignarMotorizadoCommand("PED001", "MOTXXX", "USR001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("MOTORIZADO_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerTarifa_TipoDiaLV_RetornaTarifaCorrectamente()
    {
        var motorizado = Inforest.Domain.Entities.Motorizado.Motorizado.Crear("MOT001", "12345678", 12m, 18m, 25m);

        var mockRepo = new Mock<IMotorizadoRepository>();
        mockRepo.Setup(r => r.ObtenerPorCodigoAsync("MOT001", It.IsAny<CancellationToken>()))
                .ReturnsAsync(motorizado);

        var handler = new ObtenerTarifaMotorizadoHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new ObtenerTarifaMotorizadoQuery("MOT001", TipoDiaTarifa.LunesViernes));

        Assert.True(result.EsExitoso);
        Assert.Equal(12m, result.Valor);
    }

    [Fact]
    public async Task RegistrarLlegada_LlamaaRepo()
    {
        var mockRepo = new Mock<IMotorizadoRepository>();
        mockRepo.Setup(r => r.RegistrarLlegadaAsync("PED001", "MOT001", It.IsAny<CancellationToken>()))
                .Returns(Task.CompletedTask);

        var handler = new RegistrarLlegadaSalidaHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new RegistrarLlegadaSalidaCommand("PED001", "MOT001", EsLlegada: true));

        Assert.True(result.EsExitoso);
        mockRepo.Verify(r => r.RegistrarLlegadaAsync("PED001", "MOT001", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task RegistrarSalida_LlamaaRepo()
    {
        var mockRepo = new Mock<IMotorizadoRepository>();
        mockRepo.Setup(r => r.RegistrarSalidaAsync("PED001", "MOT001", It.IsAny<CancellationToken>()))
                .Returns(Task.CompletedTask);

        var handler = new RegistrarLlegadaSalidaHandler(mockRepo.Object);
        var result = await handler.HandleAsync(new RegistrarLlegadaSalidaCommand("PED001", "MOT001", EsLlegada: false));

        Assert.True(result.EsExitoso);
        mockRepo.Verify(r => r.RegistrarSalidaAsync("PED001", "MOT001", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ObtenerPedidosReasignacion_RangoInvalido_RetornaFallo()
    {
        var mockRepo = new Mock<IPedidoDeliveryRepository>();
        var handler = new ObtenerPedidosReasignacionHandler(mockRepo.Object);

        var result = await handler.HandleAsync(
            new ObtenerPedidosReasignacionQuery(DateTime.Today, DateTime.Today.AddDays(-1), "PED"));

        Assert.False(result.EsExitoso);
        Assert.Equal("REASIGNACION_RANGO_FECHAS_INVALIDO", result.CodigoError);
    }

    [Fact]
    public async Task AsignarReasignacion_SuperaMaximoSinAutorizacion_RetornaFallo()
    {
        var motorizado = Inforest.Domain.Entities.Motorizado.Motorizado.Crear("MOT001", "12345678", 10m, 15m, 20m);
        var pedido = Inforest.Domain.Entities.Delivery.PedidoDelivery.Crear(
            "PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);
        var configuracion = ConfiguracionSistema.Cargar(
            tIdentificacionTributaria: null, tRazonSocial: null, tRazonComercial: null, tDireccion: null,
            tTelefono: null, tEmail: null, tWebPage: null, tMonedaN: null, tMonN: null, tMonedaE: null,
            tMonE: null, nTiempo: null, nChkTiempo: null, Impuesto1: 0, Impuesto2: 0, Impuesto3: 0,
            tImpuesto1: null, tImpuesto2: null, tImpuesto3: null, nCorrelativo: null, nDelivery: 0,
            nLlevar: 0, nCanal4: 0, nCanal5: 0, tPie: null, lBotonTrans: false, tElimina: null, tPassword: null,
            nItem: null, lLongitud: false, nLongitud: null, lPrinter: false, lAlmacen: false, lRapido: false,
            tBoton1: null, tBoton2: null, tBoton3: null, tBoton4: null, tBoton5: null, tPiePreCuenta: null,
            lInfhotel: false, tClub: null, nPunto: 0, lCierre: false, nDecimal: null, nDias: null,
            lEquivalencia: false, nCabecera: null, nDetalle: null, tVersion: null, lComboGeneral: false,
            nDiasDelivery: null, nTiempoMinutoCD: null, lMultilocal: false, lKDS: false, tOrderInfo: null,
            tOrderStatus: null, tBump: null, lDiaContableAutomatico: false, tHoraCierreDiaContable: null,
            lDiaContableManual: false, lClub: false, lImprimeDiaContable: false, nItemGuia: null,
            nCabeceraGuia: null, nDetalleGuia: null, nAsignacionMotorizado: 15, tTarifaActualMotorizado: "Tarifa Lun - Sab",
            tUsuarioTarifa: null, fRegistroTarifa: null, lEnvioChef: false, tContribuyenteEspecial: null,
            fContribuyenteEspecial: null, tDireccion2: null, lMobileUnidadNegocio: false,
            lMobilePasswordCCaja: false, lActivaConsultaDescargo: false, nCabeceraV: null, nItemV: null, nPieV: null,
            lFacturacionE: false, lControlUsuario: false, lHoraEntregaDelivery: false, tCodigoFE: null,
            tPieDocumento1: null, lAmbienteFE: false, lHuellaDigital: false, lHuellaSecugen: false,
            tAgenteRetencion: null, tFax: null, lLogCajaRapida: false, nTiempoToleranciaAnf: null,
            lEmailConfirmacion: false, tEmailConfirmacion: null, lEmailRecordatorio: false, tEmailRecordatorio: null,
            lEmailAgradecimiento: false, tEmailAgradecimiento: null, lImprimeCodigoBarras: false, lEnvioAutomatico: false,
            lControlEnviosProduccion: false, lActivaTarjeta: false, lEventos: false, lFEOfisis: false,
            tCodigoEmpresa: null, tCodigoTienda: null, tCodigoMarca: null, tCodigoUbigeo: null, lPagoAntesImpresion: false,
            lCambioContrasena: false, tCajaMesa247: null, tAdicionMesa247: null, tServidorFE: null, tBDFE: null,
            lActivaCuenCorrienteAut: false, lInteSAP: false, tServidorSAP: null, tBDSAP: null, tCodAlmcSAP: null,
            tCuentaContableCort: null, lVisor8: false, lvisortactil: false, lvisortiempo: null, lGlosaTransGratuita: null,
            lActivaCover: false, tMontoMinCover: null, tCodItemCover: null, lNcOfisis: false, tRutaFE: null,
            lCodigoQrFE: false, lactivaFechaNC: false, lCheffFiltroSalon: false, lCheffFiltroSubGrupo: false,
            lFEpape: false, lDesPagoCheque: false, lDesPagoOtro: false, tUsuarioFE: null, tClaveFE: null,
            lFESpring: false, lFECarbajal: false, lAnula: false, paramCarvajal: null, tCarvajalCorreos: null,
            lParcialNC: false, lNCElimina: false, lValidaDNI: false, lVerTCImp: false, lFETCI: false,
            lDesactivaNCFP: false, lFEBiz: false, tCodAnticipo: null, lActivaAnticipo: false, lFEGood: false,
            tMaxMotorizado: null, lStockDescargo: false, lFEubl21: false, lBloqInafecto: false, lDescargoInsumo: false,
            lEstupendoFE: false, lFEGesa: false, tMonedaE2: null, tMonE2: null, tMonedaE3: null, tMonE3: null);

        var mockMotRepo = new Mock<IMotorizadoRepository>();
        mockMotRepo.Setup(r => r.ObtenerPorCodigoAsync("MOT001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(motorizado);

        var mockPedRepo = new Mock<IPedidoDeliveryRepository>();
        mockPedRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);
        mockPedRepo.Setup(r => r.ContarAsignacionesPrincipalesMotorizadoAsync("MOT001", It.IsAny<DateTime>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(1);
        mockPedRepo.Setup(r => r.ContarReasignacionesMotorizadoAsync("MOT001", It.IsAny<DateTime>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(1);

        var mockParametroRepo = new Mock<IParametroRepository>();
        mockParametroRepo.Setup(r => r.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(configuracion);

        var handler = new AsignarReasignacionMotorizadoHandler(mockMotRepo.Object, mockPedRepo.Object, mockParametroRepo.Object);
        var result = await handler.HandleAsync(new AsignarReasignacionMotorizadoCommand("PED001", "MOT001", "USR001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("REASIGNACION_SUPERA_MONTO_MAXIMO", result.CodigoError);
        mockPedRepo.Verify(
            r => r.ActualizarReasignacionMotorizadoAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<decimal>(), It.IsAny<bool>(), It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public async Task DesasignarReasignacion_LlamaRepo()
    {
        var pedido = Inforest.Domain.Entities.Delivery.PedidoDelivery.Crear(
            "PED001", "TUR001", "USR001", "CAJA001", "DEL001", null, null);

        var mockPedRepo = new Mock<IPedidoDeliveryRepository>();
        mockPedRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);
        mockPedRepo.Setup(r => r.LimpiarReasignacionMotorizadoAsync("PED001", It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var handler = new DesasignarReasignacionMotorizadoHandler(mockPedRepo.Object);
        var result = await handler.HandleAsync(new DesasignarReasignacionMotorizadoCommand("PED001"));

        Assert.True(result.EsExitoso);
        mockPedRepo.Verify(r => r.LimpiarReasignacionMotorizadoAsync("PED001", It.IsAny<CancellationToken>()), Times.Once);
    }
}
