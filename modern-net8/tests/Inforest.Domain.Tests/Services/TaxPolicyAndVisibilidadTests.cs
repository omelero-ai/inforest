using Inforest.Domain.Entities.Configuracion;
using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Services;

namespace Inforest.Domain.Tests.Services;

/// <summary>
/// Pruebas de la política de impuestos.
/// Legacy: BR-002 — TPRODUCTO flags lImpuesto1..15 + TPARAMETRO Impuesto1/2/3.
/// </summary>
public sealed class TaxPolicyTests
{
    private static ConfiguracionSistema BuildConfig(double imp1 = 18.0, double imp2 = 0, double imp3 = 0)
        => ConfiguracionSistema.Cargar(
            tIdentificacionTributaria: null, tRazonSocial: null, tRazonComercial: null, tDireccion: null,
            tTelefono: null, tEmail: null, tWebPage: null, tMonedaN: null, tMonN: null, tMonedaE: null,
            tMonE: null, nTiempo: null, nChkTiempo: null,
            Impuesto1: imp1, Impuesto2: imp2, Impuesto3: imp3,
            tImpuesto1: imp1 > 0 ? "IGV" : null, tImpuesto2: null, tImpuesto3: null,
            nCorrelativo: null, nDelivery: 0, nLlevar: 0, nCanal4: 0, nCanal5: 0,
            tPie: null, lBotonTrans: false, tElimina: null, tPassword: null,
            nItem: null, lLongitud: false, nLongitud: null, lPrinter: false, lAlmacen: false, lRapido: false,
            tBoton1: null, tBoton2: null, tBoton3: null, tBoton4: null, tBoton5: null, tPiePreCuenta: null,
            lInfhotel: false, tClub: null, nPunto: 0, lCierre: false, nDecimal: null, nDias: null,
            lEquivalencia: false, nCabecera: null, nDetalle: null, tVersion: null, lComboGeneral: false,
            nDiasDelivery: null, nTiempoMinutoCD: null, lMultilocal: false, lKDS: false, tOrderInfo: null,
            tOrderStatus: null, tBump: null, lDiaContableAutomatico: false, tHoraCierreDiaContable: null,
            lDiaContableManual: false, lClub: false, lImprimeDiaContable: false, nItemGuia: null,
            nCabeceraGuia: null, nDetalleGuia: null, nAsignacionMotorizado: 0, tTarifaActualMotorizado: null,
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

    private static ProductoMaestro BuildProducto(bool imp1 = true, bool imp2 = false, bool imp3 = false)
        => ProductoMaestro.Crear(
            "P001", "BEB", "Gaseosa", "tester", 100m,
            impuestos: [imp1, imp2, imp3, false, false, false, false, false, false, false, false, false, false, false, false]);

    [Fact]
    public void CalcularImpuesto1_SiFlagActivo_RetornaMontoCorrespondiente()
    {
        var config = BuildConfig(imp1: 18.0);
        var producto = BuildProducto(imp1: true);

        var resultado = TaxPolicy.CalcularImpuesto1(producto, config, 118m);

        Assert.Equal(21.24m, resultado); // 118 * 18/100 = 21.24
    }

    [Fact]
    public void CalcularImpuesto1_SiFlagInactivo_RetornaCero()
    {
        var config = BuildConfig(imp1: 18.0);
        var producto = BuildProducto(imp1: false);

        var resultado = TaxPolicy.CalcularImpuesto1(producto, config, 118m);

        Assert.Equal(0m, resultado);
    }

    [Fact]
    public void CalcularPrecioNeto_DividePorFactorImpuesto()
    {
        var precioNeto = TaxPolicy.CalcularPrecioNeto(118m, 18m, 0m, 0m);

        Assert.Equal(100m, precioNeto);
    }

    [Fact]
    public void CalcularPrecioNeto_SinImpuesto_RetornaMismoPrecio()
    {
        var precioNeto = TaxPolicy.CalcularPrecioNeto(100m, 0m, 0m, 0m);

        Assert.Equal(100m, precioNeto);
    }

    [Fact]
    public void ObtenerPorcentajeTotalImpuestos_SumaPorcentajesActivos()
    {
        var config = BuildConfig(imp1: 18.0, imp2: 5.0, imp3: 0);
        var producto = BuildProducto(imp1: true, imp2: true, imp3: false);

        var total = TaxPolicy.ObtenerPorcentajeTotalImpuestos(producto, config);

        Assert.Equal(23m, total);
    }
}

/// <summary>
/// Pruebas del servicio de visibilidad de productos por canal.
/// Legacy: BR-014 — TPRODUCTO flags lLocal, lDelivery, lLlevar, lCanal4, lCanal5.
/// </summary>
public sealed class ProductoVisibilidadServiceTests
{
    private readonly ProductoVisibilidadService _service = new();

    private static ProductoMaestro BuildProducto(
        bool local = true, bool delivery = false, bool llevar = false,
        bool canal4 = false, bool canal5 = false, bool activo = true)
        => ProductoMaestro.Crear(
            "P001", "GRP", "Producto", "tester", 10m,
            local: local, delivery: delivery, llevar: llevar,
            canal4: canal4, canal5: canal5, activo: activo);

    [Fact]
    public void EsVisibleParaCanal_Local_RetornaFlagLocal()
    {
        var producto = BuildProducto(local: true, delivery: false);
        Assert.True(_service.EsVisibleParaCanal(producto, CanalVenta.Local));
        Assert.False(_service.EsVisibleParaCanal(producto, CanalVenta.Delivery));
    }

    [Fact]
    public void EsVisibleParaCanal_ProductoInactivo_RetornaFalse()
    {
        var producto = BuildProducto(local: true, activo: false);
        Assert.False(_service.EsVisibleParaCanal(producto, CanalVenta.Local));
    }

    [Fact]
    public void FiltrarPorCanal_RetornaSoloProductosVisibles()
    {
        var productos = new[]
        {
            BuildProducto(local: true,  delivery: false),
            BuildProducto(local: false, delivery: true),
            BuildProducto(local: true,  delivery: true),
        };

        var resultado = _service.FiltrarPorCanal(productos, CanalVenta.Local);

        Assert.Equal(2, resultado.Count);
    }
}

