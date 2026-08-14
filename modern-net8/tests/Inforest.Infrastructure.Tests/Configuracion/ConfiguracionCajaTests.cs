using Inforest.Domain.Entities.Configuracion;
using Inforest.Domain.Exceptions;

namespace Inforest.Infrastructure.Tests.Configuracion;

public sealed class ConfiguracionCajaTests
{
    [Fact]
    public void ConfiguracionCaja_Cargar_ConCodigoCajaVacio_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            CrearConfiguracionCaja(string.Empty));

        Assert.Equal("CONFIGURACION_CAJA_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void ConfiguracionCaja_Cargar_MapeaCorrectamente_FlagsBooleanos()
    {
        var configuracion = CrearConfiguracionCaja(
            "001",
            lComanda: true,
            lActivo: true,
            lPagoRapido: true,
            EAN13: true,
            lCajaContingencia: true);

        Assert.True(configuracion.lComanda);
        Assert.True(configuracion.lActivo);
        Assert.True(configuracion.lPagoRapido);
        Assert.True(configuracion.EAN13);
        Assert.True(configuracion.lCajaContingencia);
    }

    private static ConfiguracionCaja CrearConfiguracionCaja(
        string codigoCaja,
        bool lComanda = false,
        bool lActivo = false,
        bool lPagoRapido = false,
        bool EAN13 = false,
        bool lCajaContingencia = false)
        => ConfiguracionCaja.Cargar(
            codigoCaja,
            tDescripcion: null,
            tPrecuenta: null,
            lComanda: lComanda,
            vComanda: false,
            lMotivoEliminaC: false,
            lMotivoElimina: false,
            lActivo: lActivo,
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
            lMCPV: false,
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
            lPagoRapido: lPagoRapido,
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
            EAN13: EAN13,
            lObservacionCabDoc: false,
            tCompMesa247: null,
            lMotivoDescuento: false,
            lCajaContingencia: lCajaContingencia,
            lImpPropina: false,
            lImpcomandaf2: false,
            lPassOtrosPagos: false,
            lImpClienteCab: false,
            tImpFiscal: null);
}
