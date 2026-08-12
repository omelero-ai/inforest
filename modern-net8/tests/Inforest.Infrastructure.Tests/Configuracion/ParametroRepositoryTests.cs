using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Configuracion;
using Inforest.Infrastructure.Configuracion;
using Microsoft.Data.SqlClient;
using Moq;

namespace Inforest.Infrastructure.Tests.Configuracion;

public sealed class ParametroRepositoryTests
{
    [Fact]
    public async Task ObtenerConfiguracionAsync_RetornaNull_CuandoNoHayDatos()
    {
        var table = new DataTable();
        table.Columns.Add("tIdentificacionTributaria", typeof(string));

        var connectionFactory = new Mock<IDbConnectionFactory>();
        connectionFactory
            .Setup(factory => factory.CreateOpenConnectionAsync("Inforest", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new FakeDbConnection(table.CreateDataReader()));

        var repository = new ParametroRepository(connectionFactory.Object);

        var result = await repository.ObtenerConfiguracionAsync();

        Assert.Null(result);
    }

    [Fact]
    public void ConfiguracionSistema_SeConstruyeCorrectamente_ConFactoryMethodCargar()
    {
        var fechaRegistro = new DateTime(2026, 8, 12, 10, 30, 0, DateTimeKind.Utc);
        var configuracion = ConfiguracionSistema.Cargar(
            tIdentificacionTributaria: "20123456789",
            tRazonSocial: "Inforest SAC",
            tRazonComercial: null,
            tDireccion: null,
            tTelefono: null,
            tEmail: null,
            tWebPage: null,
            tMonedaN: null,
            tMonN: null,
            tMonedaE: null,
            tMonE: null,
            nTiempo: null,
            nChkTiempo: null,
            Impuesto1: 0d,
            Impuesto2: 0d,
            Impuesto3: 0d,
            tImpuesto1: null,
            tImpuesto2: null,
            tImpuesto3: null,
            nCorrelativo: null,
            nDelivery: 7.5d,
            nLlevar: 0d,
            nCanal4: 0d,
            nCanal5: 0d,
            tPie: null,
            lBotonTrans: false,
            tElimina: null,
            tPassword: null,
            nItem: null,
            lLongitud: false,
            nLongitud: null,
            lPrinter: false,
            lAlmacen: false,
            lRapido: false,
            tBoton1: null,
            tBoton2: null,
            tBoton3: null,
            tBoton4: null,
            tBoton5: null,
            tPiePreCuenta: null,
            lInfhotel: false,
            tClub: null,
            nPunto: 0d,
            lCierre: false,
            nDecimal: null,
            nDias: null,
            lEquivalencia: false,
            nCabecera: null,
            nDetalle: null,
            tVersion: null,
            lComboGeneral: false,
            nDiasDelivery: null,
            nTiempoMinutoCD: null,
            lMultilocal: false,
            lKDS: true,
            tOrderInfo: null,
            tOrderStatus: null,
            tBump: null,
            lDiaContableAutomatico: false,
            tHoraCierreDiaContable: null,
            lDiaContableManual: false,
            lClub: false,
            lImprimeDiaContable: false,
            nItemGuia: null,
            nCabeceraGuia: null,
            nDetalleGuia: null,
            nAsignacionMotorizado: 0d,
            tTarifaActualMotorizado: null,
            tUsuarioTarifa: null,
            fRegistroTarifa: fechaRegistro,
            lEnvioChef: false,
            tContribuyenteEspecial: null,
            fContribuyenteEspecial: null,
            tDireccion2: null,
            lMobileUnidadNegocio: false,
            lMobilePasswordCCaja: false,
            lActivaConsultaDescargo: false,
            nCabeceraV: null,
            nItemV: null,
            nPieV: null,
            lFacturacionE: false,
            lControlUsuario: false,
            lHoraEntregaDelivery: false,
            tCodigoFE: null,
            tPieDocumento1: null,
            lAmbienteFE: false,
            lHuellaDigital: false,
            lHuellaSecugen: false,
            tAgenteRetencion: null,
            tFax: null,
            lLogCajaRapida: false,
            nTiempoToleranciaAnf: null,
            lEmailConfirmacion: false,
            tEmailConfirmacion: null,
            lEmailRecordatorio: false,
            tEmailRecordatorio: null,
            lEmailAgradecimiento: false,
            tEmailAgradecimiento: null,
            lImprimeCodigoBarras: false,
            lEnvioAutomatico: false,
            lControlEnviosProduccion: false,
            lActivaTarjeta: false,
            lEventos: false,
            lFEOfisis: false,
            tCodigoEmpresa: null,
            tCodigoTienda: null,
            tCodigoMarca: null,
            tCodigoUbigeo: null,
            lPagoAntesImpresion: false,
            lCambioContrasena: false,
            tCajaMesa247: null,
            tAdicionMesa247: null,
            tServidorFE: null,
            tBDFE: null,
            lActivaCuenCorrienteAut: false,
            lInteSAP: false,
            tServidorSAP: null,
            tBDSAP: null,
            tCodAlmcSAP: null,
            tCuentaContableCort: null,
            lVisor8: false,
            lvisortactil: false,
            lvisortiempo: null,
            lGlosaTransGratuita: null,
            lActivaCover: false,
            tMontoMinCover: null,
            tCodItemCover: null,
            lNcOfisis: false,
            tRutaFE: null,
            lCodigoQrFE: false,
            lactivaFechaNC: false,
            lCheffFiltroSalon: false,
            lCheffFiltroSubGrupo: false,
            lFEpape: false,
            lDesPagoCheque: false,
            lDesPagoOtro: false,
            tUsuarioFE: null,
            tClaveFE: null,
            lFESpring: false,
            lFECarbajal: false,
            lAnula: false,
            paramCarvajal: null,
            tCarvajalCorreos: null,
            lParcialNC: false,
            lNCElimina: false,
            lValidaDNI: false,
            lVerTCImp: false,
            lFETCI: false,
            lDesactivaNCFP: false,
            lFEBiz: false,
            tCodAnticipo: null,
            lActivaAnticipo: false,
            lFEGood: false,
            tMaxMotorizado: null,
            lStockDescargo: false,
            lFEubl21: false,
            lBloqInafecto: false,
            lDescargoInsumo: false,
            lEstupendoFE: false,
            lFEGesa: false,
            tMonedaE2: null,
            tMonE2: null,
            tMonedaE3: null,
            tMonE3: null);

        Assert.Equal("20123456789", configuracion.tIdentificacionTributaria);
        Assert.Equal("Inforest SAC", configuracion.tRazonSocial);
        Assert.True(configuracion.lKDS);
        Assert.Equal(7.5d, configuracion.nDelivery);
        Assert.Equal(fechaRegistro, configuracion.fRegistroTarifa);
    }

    [Fact]
    public async Task ParametroService_ObtenerBoolAsync_RetornaValorDefecto_CuandoCacheEsNull()
    {
        var repository = new Mock<IParametroRepository>();
        repository
            .Setup(repo => repo.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync((ConfiguracionSistema?)null);

        var service = new ParametroService(repository.Object);

        var result = await service.ObtenerBoolAsync("lKDS", true);

        Assert.True(result);
    }

    [Fact]
    public async Task ParametroService_InvalidarCache_BorraElCache()
    {
        var repository = new Mock<IParametroRepository>();
        repository
            .SetupSequence(repo => repo.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(CrearConfiguracionSistema(tRazonSocial: "Primera"))
            .ReturnsAsync(CrearConfiguracionSistema(tRazonSocial: "Segunda"));

        var service = new ParametroService(repository.Object);

        var primero = await service.ObtenerAsync("tRazonSocial");
        service.InvalidarCache();
        var segundo = await service.ObtenerAsync("tRazonSocial");

        Assert.Equal("Primera", primero);
        Assert.Equal("Segunda", segundo);
        repository.Verify(repo => repo.ObtenerConfiguracionAsync(It.IsAny<CancellationToken>()), Times.Exactly(2));
    }

    private static ConfiguracionSistema CrearConfiguracionSistema(
        string? tIdentificacionTributaria = null,
        string? tRazonSocial = null,
        bool lKDS = false,
        double nDelivery = 0d,
        DateTime? fRegistroTarifa = null)
        => ConfiguracionSistema.Cargar(
            tIdentificacionTributaria: tIdentificacionTributaria,
            tRazonSocial: tRazonSocial,
            tRazonComercial: null,
            tDireccion: null,
            tTelefono: null,
            tEmail: null,
            tWebPage: null,
            tMonedaN: null,
            tMonN: null,
            tMonedaE: null,
            tMonE: null,
            nTiempo: null,
            nChkTiempo: null,
            Impuesto1: 0d,
            Impuesto2: 0d,
            Impuesto3: 0d,
            tImpuesto1: null,
            tImpuesto2: null,
            tImpuesto3: null,
            nCorrelativo: null,
            nDelivery: nDelivery,
            nLlevar: 0d,
            nCanal4: 0d,
            nCanal5: 0d,
            tPie: null,
            lBotonTrans: false,
            tElimina: null,
            tPassword: null,
            nItem: null,
            lLongitud: false,
            nLongitud: null,
            lPrinter: false,
            lAlmacen: false,
            lRapido: false,
            tBoton1: null,
            tBoton2: null,
            tBoton3: null,
            tBoton4: null,
            tBoton5: null,
            tPiePreCuenta: null,
            lInfhotel: false,
            tClub: null,
            nPunto: 0d,
            lCierre: false,
            nDecimal: null,
            nDias: null,
            lEquivalencia: false,
            nCabecera: null,
            nDetalle: null,
            tVersion: null,
            lComboGeneral: false,
            nDiasDelivery: null,
            nTiempoMinutoCD: null,
            lMultilocal: false,
            lKDS: lKDS,
            tOrderInfo: null,
            tOrderStatus: null,
            tBump: null,
            lDiaContableAutomatico: false,
            tHoraCierreDiaContable: null,
            lDiaContableManual: false,
            lClub: false,
            lImprimeDiaContable: false,
            nItemGuia: null,
            nCabeceraGuia: null,
            nDetalleGuia: null,
            nAsignacionMotorizado: 0d,
            tTarifaActualMotorizado: null,
            tUsuarioTarifa: null,
            fRegistroTarifa: fRegistroTarifa,
            lEnvioChef: false,
            tContribuyenteEspecial: null,
            fContribuyenteEspecial: null,
            tDireccion2: null,
            lMobileUnidadNegocio: false,
            lMobilePasswordCCaja: false,
            lActivaConsultaDescargo: false,
            nCabeceraV: null,
            nItemV: null,
            nPieV: null,
            lFacturacionE: false,
            lControlUsuario: false,
            lHoraEntregaDelivery: false,
            tCodigoFE: null,
            tPieDocumento1: null,
            lAmbienteFE: false,
            lHuellaDigital: false,
            lHuellaSecugen: false,
            tAgenteRetencion: null,
            tFax: null,
            lLogCajaRapida: false,
            nTiempoToleranciaAnf: null,
            lEmailConfirmacion: false,
            tEmailConfirmacion: null,
            lEmailRecordatorio: false,
            tEmailRecordatorio: null,
            lEmailAgradecimiento: false,
            tEmailAgradecimiento: null,
            lImprimeCodigoBarras: false,
            lEnvioAutomatico: false,
            lControlEnviosProduccion: false,
            lActivaTarjeta: false,
            lEventos: false,
            lFEOfisis: false,
            tCodigoEmpresa: null,
            tCodigoTienda: null,
            tCodigoMarca: null,
            tCodigoUbigeo: null,
            lPagoAntesImpresion: false,
            lCambioContrasena: false,
            tCajaMesa247: null,
            tAdicionMesa247: null,
            tServidorFE: null,
            tBDFE: null,
            lActivaCuenCorrienteAut: false,
            lInteSAP: false,
            tServidorSAP: null,
            tBDSAP: null,
            tCodAlmcSAP: null,
            tCuentaContableCort: null,
            lVisor8: false,
            lvisortactil: false,
            lvisortiempo: null,
            lGlosaTransGratuita: null,
            lActivaCover: false,
            tMontoMinCover: null,
            tCodItemCover: null,
            lNcOfisis: false,
            tRutaFE: null,
            lCodigoQrFE: false,
            lactivaFechaNC: false,
            lCheffFiltroSalon: false,
            lCheffFiltroSubGrupo: false,
            lFEpape: false,
            lDesPagoCheque: false,
            lDesPagoOtro: false,
            tUsuarioFE: null,
            tClaveFE: null,
            lFESpring: false,
            lFECarbajal: false,
            lAnula: false,
            paramCarvajal: null,
            tCarvajalCorreos: null,
            lParcialNC: false,
            lNCElimina: false,
            lValidaDNI: false,
            lVerTCImp: false,
            lFETCI: false,
            lDesactivaNCFP: false,
            lFEBiz: false,
            tCodAnticipo: null,
            lActivaAnticipo: false,
            lFEGood: false,
            tMaxMotorizado: null,
            lStockDescargo: false,
            lFEubl21: false,
            lBloqInafecto: false,
            lDescargoInsumo: false,
            lEstupendoFE: false,
            lFEGesa: false,
            tMonedaE2: null,
            tMonE2: null,
            tMonedaE3: null,
            tMonE3: null);

    private sealed class FakeDbConnection : DbConnection
    {
        private readonly DbDataReader _reader;

        public FakeDbConnection(DbDataReader reader)
        {
            _reader = reader;
        }

        private string _connectionString = string.Empty;

        [AllowNull]
        public override string ConnectionString
        {
            get => _connectionString;
            set => _connectionString = value ?? string.Empty;
        }
        public override string Database => "Inforest";
        public override string DataSource => "Fake";
        public override string ServerVersion => "1.0";
        public override ConnectionState State => ConnectionState.Open;

        public override void ChangeDatabase(string databaseName) { }
        public override void Close() { }
        public override void Open() { }
        protected override DbTransaction BeginDbTransaction(IsolationLevel isolationLevel) => new FakeDbTransaction(this, isolationLevel);
        protected override DbCommand CreateDbCommand() => new FakeDbCommand(_reader);
    }

    private sealed class FakeDbCommand : DbCommand
    {
        private readonly DbDataReader _reader;
        private readonly SqlParameterCollection _parameters = new SqlCommand().Parameters;

        public FakeDbCommand(DbDataReader reader)
        {
            _reader = reader;
        }

        private string _commandText = string.Empty;

        [AllowNull]
        public override string CommandText
        {
            get => _commandText;
            set => _commandText = value ?? string.Empty;
        }
        public override int CommandTimeout { get; set; }
        public override CommandType CommandType { get; set; }
        public override bool DesignTimeVisible { get; set; }
        public override UpdateRowSource UpdatedRowSource { get; set; }
        protected override DbConnection? DbConnection { get; set; }
        protected override DbParameterCollection DbParameterCollection => _parameters;
        protected override DbTransaction? DbTransaction { get; set; }

        public override void Cancel() { }
        public override int ExecuteNonQuery() => 1;
        public override object? ExecuteScalar() => null;
        public override void Prepare() { }
        protected override DbParameter CreateDbParameter() => new SqlParameter();
        protected override DbDataReader ExecuteDbDataReader(CommandBehavior behavior) => _reader;
        protected override Task<DbDataReader> ExecuteDbDataReaderAsync(CommandBehavior behavior, CancellationToken cancellationToken)
            => Task.FromResult(_reader);
    }
    private sealed class FakeDbTransaction : DbTransaction
    {
        private readonly DbConnection _connection;

        public FakeDbTransaction(DbConnection connection, IsolationLevel isolationLevel)
        {
            _connection = connection;
            IsolationLevel = isolationLevel;
        }

        public override IsolationLevel IsolationLevel { get; }
        protected override DbConnection DbConnection => _connection;
        public override void Commit() { }
        public override void Rollback() { }
    }

}
