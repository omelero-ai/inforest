using System.Data;
using Dapper;
using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Infrastructure.Configuracion;

/// <summary>
/// Legacy: TPARAMETRO y TCAJA se cargan en <c>modPuntoVenta.bas</c> al iniciar.
/// BR-SQL-003 y BR-SQL-004.
/// </summary>
internal sealed class ParametroRepository : IParametroRepository
{
    private const string DatabaseName = "Inforest";

    private readonly IDbConnectionFactory _connectionFactory;

    public ParametroRepository(IDbConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory;
    }

    public async Task<ConfiguracionSistema?> ObtenerConfiguracionAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(DatabaseName, ct);
        const string sql = "SELECT * FROM TPARAMETRO";

        using var reader = await connection.ExecuteReaderAsync(new CommandDefinition(sql, cancellationToken: ct));
        return reader.Read() ? MapConfiguracionSistema(reader) : null;
    }

    public async Task<IReadOnlyList<ConfiguracionCaja>> ObtenerTodasCajasAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(DatabaseName, ct);
        const string sql = "SELECT * FROM TCAJA ORDER BY tCaja";

        using var reader = await connection.ExecuteReaderAsync(new CommandDefinition(sql, cancellationToken: ct));
        var cajas = new List<ConfiguracionCaja>();

        while (reader.Read())
            cajas.Add(MapConfiguracionCaja(reader));

        return cajas;
    }

    public async Task<ConfiguracionCaja?> ObtenerConfiguracionCajaAsync(string codigoCaja, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(DatabaseName, ct);
        const string sql = "SELECT * FROM TCAJA WHERE tCaja = @CodigoCaja";

        using var reader = await connection.ExecuteReaderAsync(new CommandDefinition(
            sql,
            new { CodigoCaja = codigoCaja?.Trim() },
            cancellationToken: ct));

        return reader.Read() ? MapConfiguracionCaja(reader) : null;
    }

    private static ConfiguracionSistema MapConfiguracionSistema(IDataReader reader)
        => ConfiguracionSistema.Cargar(
            GetString(reader, "tIdentificacionTributaria"),
            GetString(reader, "tRazonSocial"),
            GetString(reader, "tRazonComercial"),
            GetString(reader, "tDireccion"),
            GetString(reader, "tTelefono"),
            GetString(reader, "tEmail"),
            GetString(reader, "tWebPage"),
            GetString(reader, "tMonedaN"),
            GetString(reader, "tMonN"),
            GetString(reader, "tMonedaE"),
            GetString(reader, "tMonE"),
            GetNullableInt(reader, "nTiempo"),
            GetNullableInt(reader, "nChkTiempo"),
            GetDouble(reader, "Impuesto1"),
            GetDouble(reader, "Impuesto2"),
            GetDouble(reader, "Impuesto3"),
            GetString(reader, "tImpuesto1"),
            GetString(reader, "tImpuesto2"),
            GetString(reader, "tImpuesto3"),
            GetNullableInt(reader, "nCorrelativo"),
            GetDouble(reader, "nDelivery"),
            GetDouble(reader, "nLlevar"),
            GetDouble(reader, "nCanal4"),
            GetDouble(reader, "nCanal5"),
            GetString(reader, "tPie"),
            GetBoolean(reader, "lBotonTrans"),
            GetString(reader, "tElimina"),
            GetString(reader, "tPassword"),
            GetNullableInt(reader, "nItem"),
            GetBoolean(reader, "lLongitud"),
            GetNullableInt(reader, "nLongitud"),
            GetBoolean(reader, "lPrinter"),
            GetBoolean(reader, "lAlmacen"),
            GetBoolean(reader, "lRapido"),
            GetString(reader, "tBoton1"),
            GetString(reader, "tBoton2"),
            GetString(reader, "tBoton3"),
            GetString(reader, "tBoton4"),
            GetString(reader, "tBoton5"),
            GetString(reader, "tPiePreCuenta"),
            GetBoolean(reader, "lInfhotel"),
            GetString(reader, "tClub"),
            GetDouble(reader, "nPunto"),
            GetBoolean(reader, "lCierre"),
            GetNullableInt(reader, "nDecimal"),
            GetNullableInt(reader, "nDias"),
            GetBoolean(reader, "lEquivalencia"),
            GetNullableInt(reader, "nCabecera"),
            GetNullableInt(reader, "nDetalle"),
            GetString(reader, "tVersion"),
            GetBoolean(reader, "lComboGeneral"),
            GetNullableInt(reader, "nDiasDelivery"),
            GetNullableInt(reader, "nTiempoMinutoCD"),
            GetBoolean(reader, "lMultilocal"),
            GetBoolean(reader, "lKDS"),
            GetString(reader, "tOrderInfo"),
            GetString(reader, "tOrderStatus"),
            GetString(reader, "tBump"),
            GetBoolean(reader, "lDiaContableAutomatico"),
            GetString(reader, "tHoraCierreDiaContable"),
            GetBoolean(reader, "lDiaContableManual"),
            GetBoolean(reader, "lClub"),
            GetBoolean(reader, "lImprimeDiaContable"),
            GetNullableInt(reader, "nItemGuia"),
            GetNullableInt(reader, "nCabeceraGuia"),
            GetNullableInt(reader, "nDetalleGuia"),
            GetDouble(reader, "nAsignacionMotorizado"),
            GetString(reader, "tTarifaActualMotorizado"),
            GetString(reader, "tUsuarioTarifa"),
            GetNullableDateTime(reader, "fRegistroTarifa"),
            GetBoolean(reader, "lEnvioChef"),
            GetString(reader, "tContribuyenteEspecial"),
            GetNullableDateTime(reader, "fContribuyenteEspecial"),
            GetString(reader, "tDireccion2"),
            GetBoolean(reader, "lMobileUnidadNegocio"),
            GetBoolean(reader, "lMobilePasswordCCaja"),
            GetBoolean(reader, "lActivaConsultaDescargo"),
            GetNullableInt(reader, "nCabeceraV"),
            GetNullableInt(reader, "nItemV"),
            GetNullableInt(reader, "nPieV"),
            GetBoolean(reader, "lFacturacionE"),
            GetBoolean(reader, "lControlUsuario"),
            GetBoolean(reader, "lHoraEntregaDelivery"),
            GetString(reader, "tCodigoFE"),
            GetString(reader, "tPieDocumento1"),
            GetBoolean(reader, "lAmbienteFE"),
            GetBoolean(reader, "lHuellaDigital"),
            GetBoolean(reader, "lHuellaSecugen"),
            GetString(reader, "tAgenteRetencion"),
            GetString(reader, "tFax"),
            GetBoolean(reader, "lLogCajaRapida"),
            GetNullableInt(reader, "nTiempoToleranciaAnf"),
            GetBoolean(reader, "lEmailConfirmacion"),
            GetString(reader, "tEmailConfirmacion"),
            GetBoolean(reader, "lEmailRecordatorio"),
            GetString(reader, "tEmailRecordatorio"),
            GetBoolean(reader, "lEmailAgradecimiento"),
            GetString(reader, "tEmailAgradecimiento"),
            GetBoolean(reader, "lBODato"),
            GetBoolean(reader, "lImprimeCodigoBarras"),
            GetBoolean(reader, "lEnvioAutomatico"),
            GetBoolean(reader, "lControlEnviosProduccion"),
            GetBoolean(reader, "lActivaTarjeta"),
            GetBoolean(reader, "lEventos"),
            GetBoolean(reader, "lFEOfisis"),
            GetString(reader, "tCodigoEmpresa"),
            GetString(reader, "tCodigoTienda"),
            GetString(reader, "tCodigoMarca"),
            GetString(reader, "tCodigoUbigeo"),
            GetBoolean(reader, "lPagoAntesImpresion"),
            GetBoolean(reader, "lCambioContrasena"),
            GetString(reader, "tCajaMesa247"),
            GetString(reader, "tAdicionMesa247"),
            GetString(reader, "tServidorFE"),
            GetString(reader, "tBDFE"),
            GetBoolean(reader, "lActivaCuenCorrienteAut"),
            GetBoolean(reader, "lInteSAP"),
            GetString(reader, "tServidorSAP"),
            GetString(reader, "tBDSAP"),
            GetString(reader, "tCodAlmcSAP"),
            GetString(reader, "tCuentaContableCort"),
            GetBoolean(reader, "lVisor8"),
            GetBoolean(reader, "lvisortactil"),
            GetNullableInt(reader, "lvisortiempo"),
            GetString(reader, "lGlosaTransGratuita"),
            GetBoolean(reader, "lActivaCover"),
            GetString(reader, "tMontoMinCover"),
            GetString(reader, "tCodItemCover"),
            GetBoolean(reader, "lNcOfisis"),
            GetString(reader, "tRutaFE"),
            GetBoolean(reader, "lCodigoQrFE"),
            GetBoolean(reader, "lactivaFechaNC"),
            GetBoolean(reader, "lCheffFiltroSalon"),
            GetBoolean(reader, "lCheffFiltroSubGrupo"),
            GetBoolean(reader, "lFEpape"),
            GetBoolean(reader, "lDesPagoCheque"),
            GetBoolean(reader, "lDesPagoOtro"),
            GetString(reader, "tUsuarioFE"),
            GetString(reader, "tClaveFE"),
            GetBoolean(reader, "lFESpring"),
            GetBoolean(reader, "lFECarbajal"),
            GetBoolean(reader, "lAnula"),
            GetString(reader, "paramCarvajal"),
            GetString(reader, "tCarvajalCorreos"),
            GetBoolean(reader, "lParcialNC"),
            GetBoolean(reader, "lNCElimina"),
            GetBoolean(reader, "lValidaDNI"),
            GetBoolean(reader, "lVerTCImp"),
            GetBoolean(reader, "lFETCI"),
            GetBoolean(reader, "lDesactivaNCFP"),
            GetBoolean(reader, "lFEBiz"),
            GetString(reader, "tCodAnticipo"),
            GetBoolean(reader, "lActivaAnticipo"),
            GetBoolean(reader, "lFEGood"),
            GetString(reader, "tMaxMotorizado"),
            GetBoolean(reader, "lStockDescargo"),
            GetBoolean(reader, "lFEubl21"),
            GetBoolean(reader, "lBloqInafecto"),
            GetBoolean(reader, "lDescargoInsumo"),
            GetBoolean(reader, "lEstupendoFE"),
            GetBoolean(reader, "lFEGesa"),
            GetString(reader, "tMonedaE2"),
            GetString(reader, "tMonE2"),
            GetString(reader, "tMonedaE3"),
            GetString(reader, "tMonE3"));

    private static ConfiguracionCaja MapConfiguracionCaja(IDataReader reader)
        => ConfiguracionCaja.Cargar(
            GetString(reader, "tCaja") ?? string.Empty,
            GetString(reader, "tDescripcion"),
            GetString(reader, "tPrecuenta"),
            GetBoolean(reader, "lComanda"),
            GetBoolean(reader, "vComanda"),
            GetBoolean(reader, "lMotivoEliminaC"),
            GetBoolean(reader, "lMotivoElimina"),
            GetBoolean(reader, "lActivo"),
            GetBoolean(reader, "lRefresca"),
            GetBoolean(reader, "lPasswordC"),
            GetBoolean(reader, "lPassword"),
            GetString(reader, "tGrupo"),
            GetBoolean(reader, "lConsumo1"),
            GetBoolean(reader, "lConsumo2"),
            GetBoolean(reader, "lConsumo3"),
            GetBoolean(reader, "lPrecuenta"),
            GetBoolean(reader, "lAdicion"),
            GetBoolean(reader, "lPrecuentaAgrupada"),
            GetString(reader, "tTipoPedido"),
            GetBoolean(reader, "lObliga"),
            GetBoolean(reader, "lMozo"),
            GetBoolean(reader, "lObligaPrinter"),
            GetBoolean(reader, "lPax"),
            GetBoolean(reader, "lObligaCierre"),
            GetBoolean(reader, "lFiltroTipoPedido"),
            GetNullableInt(reader, "nPuerto"),
            GetString(reader, "tMensaje1"),
            GetString(reader, "tMensaje2"),
            GetBoolean(reader, "lCancelacion"),
            GetBoolean(reader, "lDirecto"),
            GetBoolean(reader, "lObligaPrecuenta"),
            GetBoolean(reader, "lComboPrecuenta"),
            GetBoolean(reader, "lComboDocumento"),
            GetBoolean(reader, "lCambioMesa"),
            GetBoolean(reader, "lVisaNet"),
            GetBoolean(reader, "lImpuestoPrecuenta"),
            GetBoolean(reader, "lDocumentoAgrupado"),
            GetBoolean(reader, "lOrden"),
            GetBoolean(reader, "lValorCortesia"),
            GetBoolean(reader, "lObservacion"),
            GetBoolean(reader, "lCajaRapida"),
            GetBoolean(reader, "lPropiedadPrecuenta"),
            GetBoolean(reader, "lPropiedadDocumento"),
            GetBoolean(reader, "lPrecioNetoPrecuenta"),
            GetNullableInt(reader, "nLimitePrecuenta"),
            GetString(reader, "tUnidadNegocio"),
            GetNullableInt(reader, "nLimiteReimpresion"),
            GetBoolean(reader, "lPasswordTransferencia"),
            GetBoolean(reader, "lCD"),
            GetBoolean(reader, "lFechaEntregaDelivery"),
            GetBoolean(reader, "lMultiCajero"),
            GetBoolean(reader, "lMCPV"),
            GetBoolean(reader, "lCCVOX"),
            GetBoolean(reader, "lMotorizado"),
            GetBoolean(reader, "lEquivaDolaPrecuenta"),
            GetString(reader, "tSubAlmacen"),
            GetBoolean(reader, "lObservacionPrecuenta"),
            GetBoolean(reader, "lObservacionDocumento"),
            GetBoolean(reader, "lPasswordImportarPedido"),
            GetBoolean(reader, "lActivaImpDscAlternativa"),
            GetBoolean(reader, "lCompatibilidadTVS"),
            GetNullableInt(reader, "nLongitudBarra"),
            GetBoolean(reader, "lPagoRapido"),
            GetBoolean(reader, "lDisgrega"),
            GetBoolean(reader, "lPasswordPorCobrar"),
            GetBoolean(reader, "lModificaTipoPedido"),
            GetString(reader, "tSucursal"),
            GetNullableInt(reader, "nBalanzaPuerto"),
            GetBoolean(reader, "lCapturaPeso"),
            GetBoolean(reader, "lPagoRapidoPV"),
            GetString(reader, "tTextoConsumo"),
            GetBoolean(reader, "lSiab"),
            GetString(reader, "tSectorVenta"),
            GetBoolean(reader, "lCajaMobile"),
            GetBoolean(reader, "lBloqueaPrecuenta"),
            GetBoolean(reader, "lRotulado"),
            GetBoolean(reader, "lMultiAreaSubGrupo"),
            GetBoolean(reader, "lMultiAreaCaja"),
            GetBoolean(reader, "lHuella"),
            GetBoolean(reader, "lImprimeImagCabPrecuenta"),
            GetBoolean(reader, "lImprimeImagPiePrecuenta"),
            GetBoolean(reader, "lAccesoDespachoPedido"),
            GetBoolean(reader, "lBuscaPedidoNumero"),
            GetBoolean(reader, "lCodigoReciboIngreso"),
            GetBoolean(reader, "lPagoRapidoMod"),
            GetBoolean(reader, "lConsumo4"),
            GetBoolean(reader, "lPrecuentaNoValorizada"),
            GetBoolean(reader, "lWebAp"),
            GetBoolean(reader, "lBuscarPedidoVisualizarGrilla"),
            GetBoolean(reader, "lBuscarPedidoFiltrarMesa"),
            GetBoolean(reader, "lClaveEnvioProduccion"),
            GetBoolean(reader, "lMesa247"),
            GetBoolean(reader, "EAN13"),
            GetBoolean(reader, "lObservacionCabDoc"),
            GetString(reader, "tCompMesa247"),
            GetBoolean(reader, "lMotivoDescuento"),
            GetBoolean(reader, "lCajaContingencia"),
            GetBoolean(reader, "lImpPropina"),
            GetBoolean(reader, "lImpcomandaf2"),
            GetBoolean(reader, "lPassOtrosPagos"),
            GetBoolean(reader, "lImpClienteCab"),
            GetString(reader, "tImpFiscal"),
            GetBoolean(reader, "lTurnoCompartido"));

    private static string? GetString(IDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? null : Convert.ToString(reader.GetValue(ordinal));
    }

    private static bool GetBoolean(IDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return !reader.IsDBNull(ordinal) && Convert.ToBoolean(reader.GetValue(ordinal));
    }

    private static double GetDouble(IDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? 0d : Convert.ToDouble(reader.GetValue(ordinal));
    }

    private static int? GetNullableInt(IDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? null : Convert.ToInt32(reader.GetValue(ordinal));
    }

    private static DateTime? GetNullableDateTime(IDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? null : Convert.ToDateTime(reader.GetValue(ordinal));
    }
}
