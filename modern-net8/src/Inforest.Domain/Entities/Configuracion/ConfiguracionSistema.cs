using System;

namespace Inforest.Domain.Entities.Configuracion;

/// <summary>
/// Legacy: TPARAMETRO — tabla maestra de configuración del sistema. BR-SQL-003
/// </summary>
public sealed record ConfiguracionSistema
{
    public string? tIdentificacionTributaria { get; init; }
    public string? tRazonSocial { get; init; }
    public string? tRazonComercial { get; init; }
    public string? tDireccion { get; init; }
    public string? tTelefono { get; init; }
    public string? tEmail { get; init; }
    public string? tWebPage { get; init; }
    public string? tMonedaN { get; init; }
    public string? tMonN { get; init; }
    public string? tMonedaE { get; init; }
    public string? tMonE { get; init; }
    public int? nTiempo { get; init; }
    public int? nChkTiempo { get; init; }
    public double Impuesto1 { get; init; }
    public double Impuesto2 { get; init; }
    public double Impuesto3 { get; init; }
    public string? tImpuesto1 { get; init; }
    public string? tImpuesto2 { get; init; }
    public string? tImpuesto3 { get; init; }
    public int? nCorrelativo { get; init; }
    public double nDelivery { get; init; }
    public double nLlevar { get; init; }
    public double nCanal4 { get; init; }
    public double nCanal5 { get; init; }
    public string? tPie { get; init; }
    public bool lBotonTrans { get; init; }
    public string? tElimina { get; init; }
    public string? tPassword { get; init; }
    public int? nItem { get; init; }
    public bool lLongitud { get; init; }
    public int? nLongitud { get; init; }
    public bool lPrinter { get; init; }
    public bool lAlmacen { get; init; }
    public bool lRapido { get; init; }
    public string? tBoton1 { get; init; }
    public string? tBoton2 { get; init; }
    public string? tBoton3 { get; init; }
    public string? tBoton4 { get; init; }
    public string? tBoton5 { get; init; }
    public string? tPiePreCuenta { get; init; }
    public bool lInfhotel { get; init; }
    public string? tClub { get; init; }
    public double nPunto { get; init; }
    public bool lCierre { get; init; }
    public int? nDecimal { get; init; }
    public int? nDias { get; init; }
    public bool lEquivalencia { get; init; }
    public int? nCabecera { get; init; }
    public int? nDetalle { get; init; }
    public string? tVersion { get; init; }
    public bool lComboGeneral { get; init; }
    public int? nDiasDelivery { get; init; }
    public int? nTiempoMinutoCD { get; init; }
    public bool lMultilocal { get; init; }
    public bool lKDS { get; init; }
    public string? tOrderInfo { get; init; }
    public string? tOrderStatus { get; init; }
    public string? tBump { get; init; }
    public bool lDiaContableAutomatico { get; init; }
    public string? tHoraCierreDiaContable { get; init; }
    public bool lDiaContableManual { get; init; }
    public bool lClub { get; init; }
    public bool lImprimeDiaContable { get; init; }
    public int? nItemGuia { get; init; }
    public int? nCabeceraGuia { get; init; }
    public int? nDetalleGuia { get; init; }
    public double nAsignacionMotorizado { get; init; }
    public string? tTarifaActualMotorizado { get; init; }
    public string? tUsuarioTarifa { get; init; }
    public DateTime? fRegistroTarifa { get; init; }
    public bool lEnvioChef { get; init; }
    public string? tContribuyenteEspecial { get; init; }
    public DateTime? fContribuyenteEspecial { get; init; }
    public string? tDireccion2 { get; init; }
    public bool lMobileUnidadNegocio { get; init; }
    public bool lMobilePasswordCCaja { get; init; }
    public bool lActivaConsultaDescargo { get; init; }
    public int? nCabeceraV { get; init; }
    public int? nItemV { get; init; }
    public int? nPieV { get; init; }
    public bool lFacturacionE { get; init; }
    public bool lControlUsuario { get; init; }
    public bool lHoraEntregaDelivery { get; init; }
    public string? tCodigoFE { get; init; }
    public string? tPieDocumento1 { get; init; }
    public bool lAmbienteFE { get; init; }
    public bool lHuellaDigital { get; init; }
    public bool lHuellaSecugen { get; init; }
    public string? tAgenteRetencion { get; init; }
    public string? tFax { get; init; }
    public bool lLogCajaRapida { get; init; }
    public int? nTiempoToleranciaAnf { get; init; }
    public bool lEmailConfirmacion { get; init; }
    public string? tEmailConfirmacion { get; init; }
    public bool lEmailRecordatorio { get; init; }
    public string? tEmailRecordatorio { get; init; }
    public bool lEmailAgradecimiento { get; init; }
    public string? tEmailAgradecimiento { get; init; }
    public bool lBODato { get; init; }
    public bool lImprimeCodigoBarras { get; init; }
    public bool lEnvioAutomatico { get; init; }
    public bool lControlEnviosProduccion { get; init; }
    public bool lActivaTarjeta { get; init; }
    public bool lEventos { get; init; }
    public bool lFEOfisis { get; init; }
    public string? tCodigoEmpresa { get; init; }
    public string? tCodigoTienda { get; init; }
    public string? tCodigoMarca { get; init; }
    public string? tCodigoUbigeo { get; init; }
    public bool lPagoAntesImpresion { get; init; }
    public bool lCambioContrasena { get; init; }
    public string? tCajaMesa247 { get; init; }
    public string? tAdicionMesa247 { get; init; }
    public string? tServidorFE { get; init; }
    public string? tBDFE { get; init; }
    public bool lActivaCuenCorrienteAut { get; init; }
    public bool lInteSAP { get; init; }
    public string? tServidorSAP { get; init; }
    public string? tBDSAP { get; init; }
    public string? tCodAlmcSAP { get; init; }
    public string? tCuentaContableCort { get; init; }
    public bool lVisor8 { get; init; }
    public bool lvisortactil { get; init; }
    public int? lvisortiempo { get; init; }
    public string? lGlosaTransGratuita { get; init; }
    public bool lActivaCover { get; init; }
    public string? tMontoMinCover { get; init; }
    public string? tCodItemCover { get; init; }
    public bool lNcOfisis { get; init; }
    public string? tRutaFE { get; init; }
    public bool lCodigoQrFE { get; init; }
    public bool lactivaFechaNC { get; init; }
    public bool lCheffFiltroSalon { get; init; }
    public bool lCheffFiltroSubGrupo { get; init; }
    public bool lFEpape { get; init; }
    public bool lDesPagoCheque { get; init; }
    public bool lDesPagoOtro { get; init; }
    public string? tUsuarioFE { get; init; }
    public string? tClaveFE { get; init; }
    public bool lFESpring { get; init; }
    public bool lFECarbajal { get; init; }
    public bool lAnula { get; init; }
    public string? paramCarvajal { get; init; }
    public string? tCarvajalCorreos { get; init; }
    public bool lParcialNC { get; init; }
    public bool lNCElimina { get; init; }
    public bool lValidaDNI { get; init; }
    public bool lVerTCImp { get; init; }
    public bool lFETCI { get; init; }
    public bool lDesactivaNCFP { get; init; }
    public bool lFEBiz { get; init; }
    public string? tCodAnticipo { get; init; }
    public bool lActivaAnticipo { get; init; }
    public bool lFEGood { get; init; }
    public string? tMaxMotorizado { get; init; }
    public bool lStockDescargo { get; init; }
    public bool lFEubl21 { get; init; }
    public bool lBloqInafecto { get; init; }
    public bool lDescargoInsumo { get; init; }
    public bool lEstupendoFE { get; init; }
    public bool lFEGesa { get; init; }
    public string? tMonedaE2 { get; init; }
    public string? tMonE2 { get; init; }
    public string? tMonedaE3 { get; init; }
    public string? tMonE3 { get; init; }

    private ConfiguracionSistema()
    {
    }

    public static ConfiguracionSistema Cargar(
        string? tIdentificacionTributaria,
        string? tRazonSocial,
        string? tRazonComercial,
        string? tDireccion,
        string? tTelefono,
        string? tEmail,
        string? tWebPage,
        string? tMonedaN,
        string? tMonN,
        string? tMonedaE,
        string? tMonE,
        int? nTiempo,
        int? nChkTiempo,
        double Impuesto1,
        double Impuesto2,
        double Impuesto3,
        string? tImpuesto1,
        string? tImpuesto2,
        string? tImpuesto3,
        int? nCorrelativo,
        double nDelivery,
        double nLlevar,
        double nCanal4,
        double nCanal5,
        string? tPie,
        bool lBotonTrans,
        string? tElimina,
        string? tPassword,
        int? nItem,
        bool lLongitud,
        int? nLongitud,
        bool lPrinter,
        bool lAlmacen,
        bool lRapido,
        string? tBoton1,
        string? tBoton2,
        string? tBoton3,
        string? tBoton4,
        string? tBoton5,
        string? tPiePreCuenta,
        bool lInfhotel,
        string? tClub,
        double nPunto,
        bool lCierre,
        int? nDecimal,
        int? nDias,
        bool lEquivalencia,
        int? nCabecera,
        int? nDetalle,
        string? tVersion,
        bool lComboGeneral,
        int? nDiasDelivery,
        int? nTiempoMinutoCD,
        bool lMultilocal,
        bool lKDS,
        string? tOrderInfo,
        string? tOrderStatus,
        string? tBump,
        bool lDiaContableAutomatico,
        string? tHoraCierreDiaContable,
        bool lDiaContableManual,
        bool lClub,
        bool lImprimeDiaContable,
        int? nItemGuia,
        int? nCabeceraGuia,
        int? nDetalleGuia,
        double nAsignacionMotorizado,
        string? tTarifaActualMotorizado,
        string? tUsuarioTarifa,
        DateTime? fRegistroTarifa,
        bool lEnvioChef,
        string? tContribuyenteEspecial,
        DateTime? fContribuyenteEspecial,
        string? tDireccion2,
        bool lMobileUnidadNegocio,
        bool lMobilePasswordCCaja,
        bool lActivaConsultaDescargo,
        int? nCabeceraV,
        int? nItemV,
        int? nPieV,
        bool lFacturacionE,
        bool lControlUsuario,
        bool lHoraEntregaDelivery,
        string? tCodigoFE,
        string? tPieDocumento1,
        bool lAmbienteFE,
        bool lHuellaDigital,
        bool lHuellaSecugen,
        string? tAgenteRetencion,
        string? tFax,
        bool lLogCajaRapida,
        int? nTiempoToleranciaAnf,
        bool lEmailConfirmacion,
        string? tEmailConfirmacion,
        bool lEmailRecordatorio,
        string? tEmailRecordatorio,
        bool lEmailAgradecimiento,
        string? tEmailAgradecimiento,
        bool lBODato,
        bool lImprimeCodigoBarras,
        bool lEnvioAutomatico,
        bool lControlEnviosProduccion,
        bool lActivaTarjeta,
        bool lEventos,
        bool lFEOfisis,
        string? tCodigoEmpresa,
        string? tCodigoTienda,
        string? tCodigoMarca,
        string? tCodigoUbigeo,
        bool lPagoAntesImpresion,
        bool lCambioContrasena,
        string? tCajaMesa247,
        string? tAdicionMesa247,
        string? tServidorFE,
        string? tBDFE,
        bool lActivaCuenCorrienteAut,
        bool lInteSAP,
        string? tServidorSAP,
        string? tBDSAP,
        string? tCodAlmcSAP,
        string? tCuentaContableCort,
        bool lVisor8,
        bool lvisortactil,
        int? lvisortiempo,
        string? lGlosaTransGratuita,
        bool lActivaCover,
        string? tMontoMinCover,
        string? tCodItemCover,
        bool lNcOfisis,
        string? tRutaFE,
        bool lCodigoQrFE,
        bool lactivaFechaNC,
        bool lCheffFiltroSalon,
        bool lCheffFiltroSubGrupo,
        bool lFEpape,
        bool lDesPagoCheque,
        bool lDesPagoOtro,
        string? tUsuarioFE,
        string? tClaveFE,
        bool lFESpring,
        bool lFECarbajal,
        bool lAnula,
        string? paramCarvajal,
        string? tCarvajalCorreos,
        bool lParcialNC,
        bool lNCElimina,
        bool lValidaDNI,
        bool lVerTCImp,
        bool lFETCI,
        bool lDesactivaNCFP,
        bool lFEBiz,
        string? tCodAnticipo,
        bool lActivaAnticipo,
        bool lFEGood,
        string? tMaxMotorizado,
        bool lStockDescargo,
        bool lFEubl21,
        bool lBloqInafecto,
        bool lDescargoInsumo,
        bool lEstupendoFE,
        bool lFEGesa,
        string? tMonedaE2,
        string? tMonE2,
        string? tMonedaE3,
        string? tMonE3)
    {
        return new ConfiguracionSistema
        {
            tIdentificacionTributaria = tIdentificacionTributaria,
            tRazonSocial = tRazonSocial,
            tRazonComercial = tRazonComercial,
            tDireccion = tDireccion,
            tTelefono = tTelefono,
            tEmail = tEmail,
            tWebPage = tWebPage,
            tMonedaN = tMonedaN,
            tMonN = tMonN,
            tMonedaE = tMonedaE,
            tMonE = tMonE,
            nTiempo = nTiempo,
            nChkTiempo = nChkTiempo,
            Impuesto1 = Impuesto1,
            Impuesto2 = Impuesto2,
            Impuesto3 = Impuesto3,
            tImpuesto1 = tImpuesto1,
            tImpuesto2 = tImpuesto2,
            tImpuesto3 = tImpuesto3,
            nCorrelativo = nCorrelativo,
            nDelivery = nDelivery,
            nLlevar = nLlevar,
            nCanal4 = nCanal4,
            nCanal5 = nCanal5,
            tPie = tPie,
            lBotonTrans = lBotonTrans,
            tElimina = tElimina,
            tPassword = tPassword,
            nItem = nItem,
            lLongitud = lLongitud,
            nLongitud = nLongitud,
            lPrinter = lPrinter,
            lAlmacen = lAlmacen,
            lRapido = lRapido,
            tBoton1 = tBoton1,
            tBoton2 = tBoton2,
            tBoton3 = tBoton3,
            tBoton4 = tBoton4,
            tBoton5 = tBoton5,
            tPiePreCuenta = tPiePreCuenta,
            lInfhotel = lInfhotel,
            tClub = tClub,
            nPunto = nPunto,
            lCierre = lCierre,
            nDecimal = nDecimal,
            nDias = nDias,
            lEquivalencia = lEquivalencia,
            nCabecera = nCabecera,
            nDetalle = nDetalle,
            tVersion = tVersion,
            lComboGeneral = lComboGeneral,
            nDiasDelivery = nDiasDelivery,
            nTiempoMinutoCD = nTiempoMinutoCD,
            lMultilocal = lMultilocal,
            lKDS = lKDS,
            tOrderInfo = tOrderInfo,
            tOrderStatus = tOrderStatus,
            tBump = tBump,
            lDiaContableAutomatico = lDiaContableAutomatico,
            tHoraCierreDiaContable = tHoraCierreDiaContable,
            lDiaContableManual = lDiaContableManual,
            lClub = lClub,
            lImprimeDiaContable = lImprimeDiaContable,
            nItemGuia = nItemGuia,
            nCabeceraGuia = nCabeceraGuia,
            nDetalleGuia = nDetalleGuia,
            nAsignacionMotorizado = nAsignacionMotorizado,
            tTarifaActualMotorizado = tTarifaActualMotorizado,
            tUsuarioTarifa = tUsuarioTarifa,
            fRegistroTarifa = fRegistroTarifa,
            lEnvioChef = lEnvioChef,
            tContribuyenteEspecial = tContribuyenteEspecial,
            fContribuyenteEspecial = fContribuyenteEspecial,
            tDireccion2 = tDireccion2,
            lMobileUnidadNegocio = lMobileUnidadNegocio,
            lMobilePasswordCCaja = lMobilePasswordCCaja,
            lActivaConsultaDescargo = lActivaConsultaDescargo,
            nCabeceraV = nCabeceraV,
            nItemV = nItemV,
            nPieV = nPieV,
            lFacturacionE = lFacturacionE,
            lControlUsuario = lControlUsuario,
            lHoraEntregaDelivery = lHoraEntregaDelivery,
            tCodigoFE = tCodigoFE,
            tPieDocumento1 = tPieDocumento1,
            lAmbienteFE = lAmbienteFE,
            lHuellaDigital = lHuellaDigital,
            lHuellaSecugen = lHuellaSecugen,
            tAgenteRetencion = tAgenteRetencion,
            tFax = tFax,
            lLogCajaRapida = lLogCajaRapida,
            nTiempoToleranciaAnf = nTiempoToleranciaAnf,
            lEmailConfirmacion = lEmailConfirmacion,
            tEmailConfirmacion = tEmailConfirmacion,
            lEmailRecordatorio = lEmailRecordatorio,
            tEmailRecordatorio = tEmailRecordatorio,
            lEmailAgradecimiento = lEmailAgradecimiento,
            tEmailAgradecimiento = tEmailAgradecimiento,
            lBODato = lBODato,
            lImprimeCodigoBarras = lImprimeCodigoBarras,
            lEnvioAutomatico = lEnvioAutomatico,
            lControlEnviosProduccion = lControlEnviosProduccion,
            lActivaTarjeta = lActivaTarjeta,
            lEventos = lEventos,
            lFEOfisis = lFEOfisis,
            tCodigoEmpresa = tCodigoEmpresa,
            tCodigoTienda = tCodigoTienda,
            tCodigoMarca = tCodigoMarca,
            tCodigoUbigeo = tCodigoUbigeo,
            lPagoAntesImpresion = lPagoAntesImpresion,
            lCambioContrasena = lCambioContrasena,
            tCajaMesa247 = tCajaMesa247,
            tAdicionMesa247 = tAdicionMesa247,
            tServidorFE = tServidorFE,
            tBDFE = tBDFE,
            lActivaCuenCorrienteAut = lActivaCuenCorrienteAut,
            lInteSAP = lInteSAP,
            tServidorSAP = tServidorSAP,
            tBDSAP = tBDSAP,
            tCodAlmcSAP = tCodAlmcSAP,
            tCuentaContableCort = tCuentaContableCort,
            lVisor8 = lVisor8,
            lvisortactil = lvisortactil,
            lvisortiempo = lvisortiempo,
            lGlosaTransGratuita = lGlosaTransGratuita,
            lActivaCover = lActivaCover,
            tMontoMinCover = tMontoMinCover,
            tCodItemCover = tCodItemCover,
            lNcOfisis = lNcOfisis,
            tRutaFE = tRutaFE,
            lCodigoQrFE = lCodigoQrFE,
            lactivaFechaNC = lactivaFechaNC,
            lCheffFiltroSalon = lCheffFiltroSalon,
            lCheffFiltroSubGrupo = lCheffFiltroSubGrupo,
            lFEpape = lFEpape,
            lDesPagoCheque = lDesPagoCheque,
            lDesPagoOtro = lDesPagoOtro,
            tUsuarioFE = tUsuarioFE,
            tClaveFE = tClaveFE,
            lFESpring = lFESpring,
            lFECarbajal = lFECarbajal,
            lAnula = lAnula,
            paramCarvajal = paramCarvajal,
            tCarvajalCorreos = tCarvajalCorreos,
            lParcialNC = lParcialNC,
            lNCElimina = lNCElimina,
            lValidaDNI = lValidaDNI,
            lVerTCImp = lVerTCImp,
            lFETCI = lFETCI,
            lDesactivaNCFP = lDesactivaNCFP,
            lFEBiz = lFEBiz,
            tCodAnticipo = tCodAnticipo,
            lActivaAnticipo = lActivaAnticipo,
            lFEGood = lFEGood,
            tMaxMotorizado = tMaxMotorizado,
            lStockDescargo = lStockDescargo,
            lFEubl21 = lFEubl21,
            lBloqInafecto = lBloqInafecto,
            lDescargoInsumo = lDescargoInsumo,
            lEstupendoFE = lEstupendoFE,
            lFEGesa = lFEGesa,
            tMonedaE2 = tMonedaE2,
            tMonE2 = tMonE2,
            tMonedaE3 = tMonedaE3,
            tMonE3 = tMonE3,
        };
    }
}
