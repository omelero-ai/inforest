using System;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Configuracion;

/// <summary>
/// Legacy: TCAJA — configuración operativa por caja. BR-SQL-004
/// </summary>
public sealed record ConfiguracionCaja
{
    public string? tCaja { get; init; }
    public string? tDescripcion { get; init; }
    public string? tPrecuenta { get; init; }
    public bool lComanda { get; init; }
    public bool vComanda { get; init; }
    public bool lMotivoEliminaC { get; init; }
    public bool lMotivoElimina { get; init; }
    public bool lActivo { get; init; }
    public bool lRefresca { get; init; }
    public bool lPasswordC { get; init; }
    public bool lPassword { get; init; }
    public string? tGrupo { get; init; }
    public bool lConsumo1 { get; init; }
    public bool lConsumo2 { get; init; }
    public bool lConsumo3 { get; init; }
    public bool lPrecuenta { get; init; }
    public bool lAdicion { get; init; }
    public bool lPrecuentaAgrupada { get; init; }
    public string? tTipoPedido { get; init; }
    public bool lObliga { get; init; }
    public bool lMozo { get; init; }
    public bool lObligaPrinter { get; init; }
    public bool lPax { get; init; }
    public bool lObligaCierre { get; init; }
    public bool lFiltroTipoPedido { get; init; }
    public int? nPuerto { get; init; }
    public string? tMensaje1 { get; init; }
    public string? tMensaje2 { get; init; }
    public bool lCancelacion { get; init; }
    public bool lDirecto { get; init; }
    public bool lObligaPrecuenta { get; init; }
    public bool lComboPrecuenta { get; init; }
    public bool lComboDocumento { get; init; }
    public bool lCambioMesa { get; init; }
    public bool lVisaNet { get; init; }
    public bool lImpuestoPrecuenta { get; init; }
    public bool lDocumentoAgrupado { get; init; }
    public bool lOrden { get; init; }
    public bool lValorCortesia { get; init; }
    public bool lObservacion { get; init; }
    public bool lCajaRapida { get; init; }
    public bool lPropiedadPrecuenta { get; init; }
    public bool lPropiedadDocumento { get; init; }
    public bool lPrecioNetoPrecuenta { get; init; }
    public int? nLimitePrecuenta { get; init; }
    public string? tUnidadNegocio { get; init; }
    public int? nLimiteReimpresion { get; init; }
    public bool lPasswordTransferencia { get; init; }
    public bool lCD { get; init; }
    public bool lFechaEntregaDelivery { get; init; }
    public bool lMultiCajero { get; init; }
    public bool lMCPV { get; init; }
    public bool lCCVOX { get; init; }
    public bool lMotorizado { get; init; }
    public bool lEquivaDolaPrecuenta { get; init; }
    public string? tSubAlmacen { get; init; }
    public bool lObservacionPrecuenta { get; init; }
    public bool lObservacionDocumento { get; init; }
    public bool lPasswordImportarPedido { get; init; }
    public bool lActivaImpDscAlternativa { get; init; }
    public bool lCompatibilidadTVS { get; init; }
    public int? nLongitudBarra { get; init; }
    public bool lPagoRapido { get; init; }
    public bool lDisgrega { get; init; }
    public bool lPasswordPorCobrar { get; init; }
    public bool lModificaTipoPedido { get; init; }
    public string? tSucursal { get; init; }
    public int? nBalanzaPuerto { get; init; }
    public bool lCapturaPeso { get; init; }
    public bool lPagoRapidoPV { get; init; }
    public string? tTextoConsumo { get; init; }
    public bool lSiab { get; init; }
    public string? tSectorVenta { get; init; }
    public bool lCajaMobile { get; init; }
    public bool lBloqueaPrecuenta { get; init; }
    public bool lRotulado { get; init; }
    public bool lMultiAreaSubGrupo { get; init; }
    public bool lMultiAreaCaja { get; init; }
    public bool lHuella { get; init; }
    public bool lImprimeImagCabPrecuenta { get; init; }
    public bool lImprimeImagPiePrecuenta { get; init; }
    public bool lAccesoDespachoPedido { get; init; }
    public bool lBuscaPedidoNumero { get; init; }
    public bool lCodigoReciboIngreso { get; init; }
    public bool lPagoRapidoMod { get; init; }
    public bool lConsumo4 { get; init; }
    public bool lPrecuentaNoValorizada { get; init; }
    public bool lWebAp { get; init; }
    public bool lBuscarPedidoVisualizarGrilla { get; init; }
    public bool lBuscarPedidoFiltrarMesa { get; init; }
    public bool lClaveEnvioProduccion { get; init; }
    public bool lMesa247 { get; init; }
    public bool EAN13 { get; init; }
    public bool lObservacionCabDoc { get; init; }
    public string? tCompMesa247 { get; init; }
    public bool lMotivoDescuento { get; init; }
    public bool lCajaContingencia { get; init; }
    public bool lImpPropina { get; init; }
    public bool lImpcomandaf2 { get; init; }
    public bool lPassOtrosPagos { get; init; }
    public bool lImpClienteCab { get; init; }
    public string? tImpFiscal { get; init; }

    private ConfiguracionCaja()
    {
    }

    public static ConfiguracionCaja Cargar(
        string? codigoCaja,
        string? tDescripcion,
        string? tPrecuenta,
        bool lComanda,
        bool vComanda,
        bool lMotivoEliminaC,
        bool lMotivoElimina,
        bool lActivo,
        bool lRefresca,
        bool lPasswordC,
        bool lPassword,
        string? tGrupo,
        bool lConsumo1,
        bool lConsumo2,
        bool lConsumo3,
        bool lPrecuenta,
        bool lAdicion,
        bool lPrecuentaAgrupada,
        string? tTipoPedido,
        bool lObliga,
        bool lMozo,
        bool lObligaPrinter,
        bool lPax,
        bool lObligaCierre,
        bool lFiltroTipoPedido,
        int? nPuerto,
        string? tMensaje1,
        string? tMensaje2,
        bool lCancelacion,
        bool lDirecto,
        bool lObligaPrecuenta,
        bool lComboPrecuenta,
        bool lComboDocumento,
        bool lCambioMesa,
        bool lVisaNet,
        bool lImpuestoPrecuenta,
        bool lDocumentoAgrupado,
        bool lOrden,
        bool lValorCortesia,
        bool lObservacion,
        bool lCajaRapida,
        bool lPropiedadPrecuenta,
        bool lPropiedadDocumento,
        bool lPrecioNetoPrecuenta,
        int? nLimitePrecuenta,
        string? tUnidadNegocio,
        int? nLimiteReimpresion,
        bool lPasswordTransferencia,
        bool lCD,
        bool lFechaEntregaDelivery,
        bool lMultiCajero,
        bool lMCPV,
        bool lCCVOX,
        bool lMotorizado,
        bool lEquivaDolaPrecuenta,
        string? tSubAlmacen,
        bool lObservacionPrecuenta,
        bool lObservacionDocumento,
        bool lPasswordImportarPedido,
        bool lActivaImpDscAlternativa,
        bool lCompatibilidadTVS,
        int? nLongitudBarra,
        bool lPagoRapido,
        bool lDisgrega,
        bool lPasswordPorCobrar,
        bool lModificaTipoPedido,
        string? tSucursal,
        int? nBalanzaPuerto,
        bool lCapturaPeso,
        bool lPagoRapidoPV,
        string? tTextoConsumo,
        bool lSiab,
        string? tSectorVenta,
        bool lCajaMobile,
        bool lBloqueaPrecuenta,
        bool lRotulado,
        bool lMultiAreaSubGrupo,
        bool lMultiAreaCaja,
        bool lHuella,
        bool lImprimeImagCabPrecuenta,
        bool lImprimeImagPiePrecuenta,
        bool lAccesoDespachoPedido,
        bool lBuscaPedidoNumero,
        bool lCodigoReciboIngreso,
        bool lPagoRapidoMod,
        bool lConsumo4,
        bool lPrecuentaNoValorizada,
        bool lWebAp,
        bool lBuscarPedidoVisualizarGrilla,
        bool lBuscarPedidoFiltrarMesa,
        bool lClaveEnvioProduccion,
        bool lMesa247,
        bool EAN13,
        bool lObservacionCabDoc,
        string? tCompMesa247,
        bool lMotivoDescuento,
        bool lCajaContingencia,
        bool lImpPropina,
        bool lImpcomandaf2,
        bool lPassOtrosPagos,
        bool lImpClienteCab,
        string? tImpFiscal)
    {
        if (string.IsNullOrWhiteSpace(codigoCaja))
            throw new DomainException("El código de caja es obligatorio.", "CONFIGURACION_CAJA_CODIGO_REQUERIDO");

        return new ConfiguracionCaja
        {
            tCaja = codigoCaja.Trim(),
            tDescripcion = tDescripcion,
            tPrecuenta = tPrecuenta,
            lComanda = lComanda,
            vComanda = vComanda,
            lMotivoEliminaC = lMotivoEliminaC,
            lMotivoElimina = lMotivoElimina,
            lActivo = lActivo,
            lRefresca = lRefresca,
            lPasswordC = lPasswordC,
            lPassword = lPassword,
            tGrupo = tGrupo,
            lConsumo1 = lConsumo1,
            lConsumo2 = lConsumo2,
            lConsumo3 = lConsumo3,
            lPrecuenta = lPrecuenta,
            lAdicion = lAdicion,
            lPrecuentaAgrupada = lPrecuentaAgrupada,
            tTipoPedido = tTipoPedido,
            lObliga = lObliga,
            lMozo = lMozo,
            lObligaPrinter = lObligaPrinter,
            lPax = lPax,
            lObligaCierre = lObligaCierre,
            lFiltroTipoPedido = lFiltroTipoPedido,
            nPuerto = nPuerto,
            tMensaje1 = tMensaje1,
            tMensaje2 = tMensaje2,
            lCancelacion = lCancelacion,
            lDirecto = lDirecto,
            lObligaPrecuenta = lObligaPrecuenta,
            lComboPrecuenta = lComboPrecuenta,
            lComboDocumento = lComboDocumento,
            lCambioMesa = lCambioMesa,
            lVisaNet = lVisaNet,
            lImpuestoPrecuenta = lImpuestoPrecuenta,
            lDocumentoAgrupado = lDocumentoAgrupado,
            lOrden = lOrden,
            lValorCortesia = lValorCortesia,
            lObservacion = lObservacion,
            lCajaRapida = lCajaRapida,
            lPropiedadPrecuenta = lPropiedadPrecuenta,
            lPropiedadDocumento = lPropiedadDocumento,
            lPrecioNetoPrecuenta = lPrecioNetoPrecuenta,
            nLimitePrecuenta = nLimitePrecuenta,
            tUnidadNegocio = tUnidadNegocio,
            nLimiteReimpresion = nLimiteReimpresion,
            lPasswordTransferencia = lPasswordTransferencia,
            lCD = lCD,
            lFechaEntregaDelivery = lFechaEntregaDelivery,
            lMultiCajero = lMultiCajero,
            lMCPV = lMCPV,
            lCCVOX = lCCVOX,
            lMotorizado = lMotorizado,
            lEquivaDolaPrecuenta = lEquivaDolaPrecuenta,
            tSubAlmacen = tSubAlmacen,
            lObservacionPrecuenta = lObservacionPrecuenta,
            lObservacionDocumento = lObservacionDocumento,
            lPasswordImportarPedido = lPasswordImportarPedido,
            lActivaImpDscAlternativa = lActivaImpDscAlternativa,
            lCompatibilidadTVS = lCompatibilidadTVS,
            nLongitudBarra = nLongitudBarra,
            lPagoRapido = lPagoRapido,
            lDisgrega = lDisgrega,
            lPasswordPorCobrar = lPasswordPorCobrar,
            lModificaTipoPedido = lModificaTipoPedido,
            tSucursal = tSucursal,
            nBalanzaPuerto = nBalanzaPuerto,
            lCapturaPeso = lCapturaPeso,
            lPagoRapidoPV = lPagoRapidoPV,
            tTextoConsumo = tTextoConsumo,
            lSiab = lSiab,
            tSectorVenta = tSectorVenta,
            lCajaMobile = lCajaMobile,
            lBloqueaPrecuenta = lBloqueaPrecuenta,
            lRotulado = lRotulado,
            lMultiAreaSubGrupo = lMultiAreaSubGrupo,
            lMultiAreaCaja = lMultiAreaCaja,
            lHuella = lHuella,
            lImprimeImagCabPrecuenta = lImprimeImagCabPrecuenta,
            lImprimeImagPiePrecuenta = lImprimeImagPiePrecuenta,
            lAccesoDespachoPedido = lAccesoDespachoPedido,
            lBuscaPedidoNumero = lBuscaPedidoNumero,
            lCodigoReciboIngreso = lCodigoReciboIngreso,
            lPagoRapidoMod = lPagoRapidoMod,
            lConsumo4 = lConsumo4,
            lPrecuentaNoValorizada = lPrecuentaNoValorizada,
            lWebAp = lWebAp,
            lBuscarPedidoVisualizarGrilla = lBuscarPedidoVisualizarGrilla,
            lBuscarPedidoFiltrarMesa = lBuscarPedidoFiltrarMesa,
            lClaveEnvioProduccion = lClaveEnvioProduccion,
            lMesa247 = lMesa247,
            EAN13 = EAN13,
            lObservacionCabDoc = lObservacionCabDoc,
            tCompMesa247 = tCompMesa247,
            lMotivoDescuento = lMotivoDescuento,
            lCajaContingencia = lCajaContingencia,
            lImpPropina = lImpPropina,
            lImpcomandaf2 = lImpcomandaf2,
            lPassOtrosPagos = lPassOtrosPagos,
            lImpClienteCab = lImpClienteCab,
            tImpFiscal = tImpFiscal,
        };
    }
}
