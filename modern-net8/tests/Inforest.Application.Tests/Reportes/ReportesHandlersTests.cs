using Inforest.Application.Reportes;
using Moq;
using Xunit;

namespace Inforest.Application.Tests.Reportes;

/// <summary>
/// Tests unitarios para los handlers de reportes.
/// Valida que los handlers llamen al repositorio con los parámetros correctos
/// y retornen el resultado esperado.
/// Reglas: BR-REP-001 … BR-REP-012
/// </summary>
public class ReportesHandlersTests
{
    private readonly Mock<IReporteRepository> _repoMock;

    public ReportesHandlersTests()
    {
        _repoMock = new Mock<IReporteRepository>();
    }

    // ── BR-REP-001 — Comanda ─────────────────────────────────────────────────

    [Fact]
    public async Task ObtenerReporteComandaHandler_LlamaRepositorioConParametrosCorrectos()
    {
        // Arrange
        var filas = new List<ComandaRow>
        {
            new() { TCodigoPedido = "P001", NombreProducto = "Lomo Saltado", NCantidad = 2 }
        };
        _repoMock.Setup(r => r.ObtenerComandaAsync(true, "Pedido", It.IsAny<DateTime>(), It.IsAny<DateTime>(), "", default))
            .ReturnsAsync(filas.AsReadOnly());

        var handler = new ObtenerReporteComandaHandler(_repoMock.Object);
        var query = new ObtenerReporteComandaQuery(
            FlagTipo: true,
            Orden: "Pedido",
            FechaInicio: new DateTime(2026, 1, 1),
            FechaFin: new DateTime(2026, 1, 31));

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.NotNull(resultado);
        Assert.Single(resultado.Filas);
        Assert.Equal("P001", resultado.Filas[0].TCodigoPedido);
        Assert.Equal("RepComandaDetallado.frx", resultado.NombrePlantilla);
        _repoMock.Verify(r => r.ObtenerComandaAsync(true, "Pedido", It.IsAny<DateTime>(), It.IsAny<DateTime>(), "", default), Times.Once);
    }

    [Fact]
    public async Task ObtenerReporteComandaHandler_ModoResumido_UsaPlantillaCorrecta()
    {
        // Arrange
        _repoMock.Setup(r => r.ObtenerComandaAsync(false, It.IsAny<string>(), It.IsAny<DateTime>(), It.IsAny<DateTime>(), It.IsAny<string>(), default))
            .ReturnsAsync(new List<ComandaRow>().AsReadOnly());

        var handler = new ObtenerReporteComandaHandler(_repoMock.Object);
        var query = new ObtenerReporteComandaQuery(false, "Fecha", DateTime.Today, DateTime.Today);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Equal("RepComanda.frx", resultado.NombrePlantilla);
    }

    // ── BR-REP-002 — Propina ─────────────────────────────────────────────────

    [Fact]
    public async Task ObtenerReportePropinaHandler_LlamaRepositorioConFechasCorrectas()
    {
        // Arrange
        var inicio = new DateTime(2026, 1, 1);
        var fin = new DateTime(2026, 1, 31);
        _repoMock.Setup(r => r.ObtenerPropinaAsync(inicio, fin, "", default))
            .ReturnsAsync(new List<PropinaRow>
            {
                new() { Propina = 5.00, Trabajador = "Juan" }
            }.AsReadOnly());

        var handler = new ObtenerReportePropinaHandler(_repoMock.Object);
        var query = new ObtenerReportePropinaQuery(inicio, fin);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Single(resultado.Filas);
        Assert.Equal(5.00, resultado.Filas[0].Propina);
        Assert.Equal("RepPropina.frx", resultado.NombrePlantilla);
    }

    [Fact]
    public async Task ObtenerReportePropinaHandler_EsResumido_RetornaPlantillaResumido()
    {
        // Arrange — BR-REP-002 modo Resumido (optOpcion(1) del Legacy)
        var inicio = new DateTime(2026, 2, 1);
        var fin = new DateTime(2026, 2, 28);
        _repoMock.Setup(r => r.ObtenerPropinaAsync(inicio, fin, "tmozo = '01'", default))
            .ReturnsAsync(new List<PropinaRow>
            {
                new() { Propina = 8.50, Trabajador = "Maria", TMozo = "01" }
            }.AsReadOnly());

        var handler = new ObtenerReportePropinaHandler(_repoMock.Object);
        var query = new ObtenerReportePropinaQuery(inicio, fin, Condicion: "tmozo = '01'", EsDetallado: false);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Single(resultado.Filas);
        Assert.Equal(8.50, resultado.Filas[0].Propina);
        Assert.Equal("RepPropinaResumido.frx", resultado.NombrePlantilla);
    }

    // ── BR-REP-003 — PrincipalCliente ────────────────────────────────────────

    [Fact]
    public async Task ObtenerReportePrincipalClienteHandler_ModoDetalle_RetornaFilasDetalle()
    {
        // Arrange
        _repoMock.Setup(r => r.ObtenerPrincipalClienteDetalleAsync(0, "", It.IsAny<DateTime>(), It.IsAny<DateTime>(), default))
            .ReturnsAsync(new List<PrincipalClienteDetalleRow>
            {
                new() { TCodigoCliente = "C001", TEmpresa = "ACME SAC", NVenta = 500 }
            }.AsReadOnly());

        var handler = new ObtenerReportePrincipalClienteHandler(_repoMock.Object);
        var query = new ObtenerReportePrincipalClienteQuery(true, 0, "", DateTime.Today.AddDays(-30), DateTime.Today);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Single(resultado.Filas);
        Assert.Contains("Detalle", resultado.NombrePlantilla);
    }

    [Fact]
    public async Task ObtenerReportePrincipalClienteHandler_ModoResumen_RetornaFilasResumen()
    {
        // Arrange
        _repoMock.Setup(r => r.ObtenerPrincipalClienteResumenAsync(100, "", It.IsAny<DateTime>(), It.IsAny<DateTime>(), default))
            .ReturnsAsync(new List<PrincipalClienteResumenRow>
            {
                new() { TCodigoCliente = "C001", TEmpresa = "ACME SAC", Cantidad = 10 }
            }.AsReadOnly());

        var handler = new ObtenerReportePrincipalClienteHandler(_repoMock.Object);
        var query = new ObtenerReportePrincipalClienteQuery(false, 100, "", DateTime.Today.AddDays(-30), DateTime.Today);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Single(resultado.Filas);
        Assert.Contains("Resumen", resultado.NombrePlantilla);
    }

    // ── BR-REP-004 — CtaCte Integrado ────────────────────────────────────────

    [Fact]
    public async Task ObtenerReporteCtaCteIntegradoHandler_FlagTipo1_UsaPlantillaConsolidado()
    {
        // Arrange
        _repoMock.Setup(r => r.ObtenerCtaCteIntegradoAsync("1", It.IsAny<DateTime>(), It.IsAny<DateTime>(), default))
            .ReturnsAsync(new List<CtaCteIntegradoRow>().AsReadOnly());

        var handler = new ObtenerReporteCtaCteIntegradoHandler(_repoMock.Object);
        var query = new ObtenerReporteCtaCteIntegradoQuery("1", DateTime.Today.AddMonths(-1), DateTime.Today);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Contains("Consolidado", resultado.NombrePlantilla);
    }

    [Fact]
    public async Task ObtenerReporteCtaCteIntegradoHandler_FlagTipo2_UsaPlantillaDetallado()
    {
        // Arrange
        _repoMock.Setup(r => r.ObtenerCtaCteIntegradoAsync("2", It.IsAny<DateTime>(), It.IsAny<DateTime>(), default))
            .ReturnsAsync(new List<CtaCteIntegradoRow>().AsReadOnly());

        var handler = new ObtenerReporteCtaCteIntegradoHandler(_repoMock.Object);
        var query = new ObtenerReporteCtaCteIntegradoQuery("2", DateTime.Today.AddMonths(-1), DateTime.Today);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Contains("Detallado", resultado.NombrePlantilla);
    }

    [Fact]
    public async Task ObtenerReporteCtaCteOperativaHandler_ModoConsolidado_UsaPlantillaCorrecta()
    {
        _repoMock.Setup(r => r.ObtenerCtaCteOperativaAsync(It.Is<CtaCteOperativaParametros>(p => p.FlagConsolidado), default))
            .ReturnsAsync(new List<CtaCteOperativaRow>().AsReadOnly());

        var handler = new ObtenerReporteCtaCteOperativaHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteCtaCteOperativaQuery(new CtaCteOperativaParametros
        {
            FlagConsolidado = true,
            FechaInicio = DateTime.Today.AddDays(-7),
            FechaFin = DateTime.Today
        }));

        Assert.Equal("RepCtaCteConsolidado.frx", resultado.NombrePlantilla);
    }

    [Fact]
    public async Task ObtenerReporteCtaCteOperativaHandler_ModoDetallado_RetornaFilas()
    {
        _repoMock.Setup(r => r.ObtenerCtaCteOperativaAsync(It.Is<CtaCteOperativaParametros>(p => p.FlagDetalle), default))
            .ReturnsAsync(new List<CtaCteOperativaRow>
            {
                new() { Descripcion = "ACME SAC", TCodigoPedido = "P001", Producto = "Lomo", NVenta = 32.5 }
            }.AsReadOnly());

        var handler = new ObtenerReporteCtaCteOperativaHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteCtaCteOperativaQuery(new CtaCteOperativaParametros
        {
            FlagDetalle = true,
            FechaInicio = DateTime.Today.AddDays(-7),
            FechaFin = DateTime.Today
        }));

        Assert.Single(resultado.Filas);
        Assert.Equal("RepCtaCteDetallado.frx", resultado.NombrePlantilla);
    }

    // ── BR-REP-SQL-DYN-001 — Expresión precio segura ─────────────────────────

    [Theory]
    [InlineData(ExpresionPrecio.Venta, "nVenta")]
    [InlineData(ExpresionPrecio.Neto, "nNeto")]
    [InlineData(ExpresionPrecio.Costo, "nInsumo")]
    public void ExpresionPrecio_ToSqlExpresion_RetornaExpresionSegura(ExpresionPrecio precio, string fragmentoEsperado)
    {
        // Act
        var expresion = precio.ToSqlExpresion();

        // Assert: la expresión nunca es vacía y contiene el fragmento esperado
        Assert.NotEmpty(expresion);
        Assert.Contains(fragmentoEsperado, expresion);
        // Garantía: no contiene delimitadores de SQL injection básicos
        Assert.DoesNotContain(";", expresion);
        Assert.DoesNotContain("--", expresion);
        Assert.DoesNotContain("'", expresion);
    }

    // ── BR-REP-009 — VentaMensualIntegrado ───────────────────────────────────

    [Fact]
    public async Task ObtenerReporteVentaMensualIntegradoHandler_TipoDiario_UsaPlantillaDiaria()
    {
        // Arrange
        _repoMock.Setup(r => r.ObtenerVentaMensualIntegradoAsync(It.IsAny<VentaMensualIntegradoParametros>(), default))
            .ReturnsAsync(new List<VentaMensualIntegradoRow>().AsReadOnly());

        var handler = new ObtenerReporteVentaMensualIntegradoHandler(_repoMock.Object);
        var parametros = new VentaMensualIntegradoParametros { Tipo = "D", Ano = "2026", Mes = "01" };
        var query = new ObtenerReporteVentaMensualIntegradoQuery(parametros);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Contains("Diaria", resultado.NombrePlantilla);
    }

    [Fact]
    public async Task ObtenerReporteVentaMensualIntegradoHandler_TipoMensual_UsaPlantillaMensual()
    {
        // Arrange
        _repoMock.Setup(r => r.ObtenerVentaMensualIntegradoAsync(It.IsAny<VentaMensualIntegradoParametros>(), default))
            .ReturnsAsync(new List<VentaMensualIntegradoRow>().AsReadOnly());

        var handler = new ObtenerReporteVentaMensualIntegradoHandler(_repoMock.Object);
        var parametros = new VentaMensualIntegradoParametros { Tipo = "M", Ano = "2026", Mes = "01" };
        var query = new ObtenerReporteVentaMensualIntegradoQuery(parametros);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Contains("Mensual", resultado.NombrePlantilla);
    }

    // ── BR-REP-011 — KDS TiempoPedido ────────────────────────────────────────

    [Fact]
    public async Task ObtenerReporteTiempoKdsPedidoHandler_RetornaFilasKds()
    {
        // Arrange
        var inicio = new DateTime(2026, 8, 1);
        var fin = new DateTime(2026, 8, 12);
        _repoMock.Setup(r => r.ObtenerTiempoKdsPedidoAsync(inicio, fin, default))
            .ReturnsAsync(new List<TiempoKdsPedidoRow>
            {
                new() { Pedido = "P001", Producto = "Lomo Saltado", TiempoPromedio = "00:08:00" }
            }.AsReadOnly());

        var handler = new ObtenerReporteTiempoKdsPedidoHandler(_repoMock.Object);
        var query = new ObtenerReporteTiempoKdsPedidoQuery(inicio, fin);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Single(resultado.Filas);
        Assert.Equal("00:08:00", resultado.Filas[0].TiempoPromedio);
        Assert.Equal("RepTiempoKdsPedido.frx", resultado.NombrePlantilla);
    }

    // ── BR-REP-012 — KDS TiempoProducto ──────────────────────────────────────

    [Fact]
    public async Task ObtenerReporteTiempoKdsProductoHandler_ConFiltros_LlamaRepositorioCorrectamente()
    {
        // Arrange
        _repoMock.Setup(r => r.ObtenerTiempoKdsProductoAsync(
            It.IsAny<DateTime>(), It.IsAny<DateTime>(), "GRP01", "SG01", "P001", default))
            .ReturnsAsync(new List<TiempoKdsProductoRow>
            {
                new() { Producto = "Pizza Margherita", TiempoPromedio = "00:12:00" }
            }.AsReadOnly());

        var handler = new ObtenerReporteTiempoKdsProductoHandler(_repoMock.Object);
        var query = new ObtenerReporteTiempoKdsProductoQuery(DateTime.Today.AddDays(-7), DateTime.Today, "GRP01", "SG01", "P001");

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Single(resultado.Filas);
        Assert.Equal("RepTiempoKdsProducto.frx", resultado.NombrePlantilla);
        _repoMock.Verify(r => r.ObtenerTiempoKdsProductoAsync(
            It.IsAny<DateTime>(), It.IsAny<DateTime>(), "GRP01", "SG01", "P001", default), Times.Once);
    }

    // ── BR-REP-014 — Anulación / Control de Transacciones ────────────────────

    [Fact]
    public async Task ObtenerReporteAnulacionHandler_ConFiltros_RetornaPlantillaCorrecta()
    {
        // Arrange
        var parametros = new AnulacionParametros
        {
            FechaInicio = DateTime.Today.AddDays(-7),
            FechaFin = DateTime.Today,
            FlagFacturados = true,
            FlagAnulados = true,
            FlagTransferidos = false,
            Criterio = string.Empty
        };

        _repoMock.Setup(r => r.ObtenerAnulacionAsync(parametros, default))
            .ReturnsAsync(new List<AnulacionRow>
            {
                new() { TCodigoPedido = "P001", TItem = "01", TCodigoProducto = "PROD01", NCantidad = 2, NVenta = 25.50, TEstadoItem = "A", TMotivoAnulacion = "ERR" }
            }.AsReadOnly());

        var handler = new ObtenerReporteAnulacionHandler(_repoMock.Object);
        var query = new ObtenerReporteAnulacionQuery(parametros);

        // Act
        var resultado = await handler.HandleAsync(query);

        // Assert
        Assert.Single(resultado.Filas);
        Assert.Equal("RepAnulacion.frx", resultado.NombrePlantilla);
        Assert.Equal("Control de Transacciones", resultado.TituloReporte);
        _repoMock.Verify(r => r.ObtenerAnulacionAsync(parametros, default), Times.Once);
    }

    [Fact]
    public async Task ObtenerReporteAnulacionHandler_SinResultados_RetornaListaVacia()
    {
        // Arrange
        var parametros = new AnulacionParametros
        {
            FechaInicio = DateTime.Today.AddDays(-1),
            FechaFin = DateTime.Today,
            FlagFacturados = false,
            FlagAnulados = true,
            FlagTransferidos = false
        };

        _repoMock.Setup(r => r.ObtenerAnulacionAsync(parametros, default))
            .ReturnsAsync(new List<AnulacionRow>().AsReadOnly());

        var handler = new ObtenerReporteAnulacionHandler(_repoMock.Object);

        // Act
        var resultado = await handler.HandleAsync(new ObtenerReporteAnulacionQuery(parametros));

        // Assert
        Assert.Empty(resultado.Filas);
        Assert.Equal("RepAnulacion.frx", resultado.NombrePlantilla);
    }

    [Fact]
    public async Task ObtenerReporteLiquidacionTicketHandler_ConTurno_RetornaPlantillaCorrecta()
    {
        var parametros = new LiquidacionTicketParametros
        {
            TodosLosTurnos = false,
            Turno = "T001",
            FechaInicio = DateTime.Today,
            FechaFin = DateTime.Today.AddDays(1).AddSeconds(-1)
        };

        _repoMock.Setup(r => r.ObtenerLiquidacionTicketAsync(parametros, default))
            .ReturnsAsync(new List<LiquidacionTicketRow>
            {
                new() { TTipoPedido = "01", NNeto = 100, NImpuesto1 = 18, NVenta = 118, NTotalPromedio = 2, Total00 = 2 }
            }.AsReadOnly());

        var handler = new ObtenerReporteLiquidacionTicketHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteLiquidacionTicketQuery(parametros));

        Assert.Single(resultado.Filas);
        Assert.Equal("RepLiquidacionTicket.frx", resultado.NombrePlantilla);
        Assert.Equal("Liquidación de Cajero por Ticketera", resultado.TituloReporte);
        _repoMock.Verify(r => r.ObtenerLiquidacionTicketAsync(parametros, default), Times.Once);
    }

    [Fact]
    public async Task ObtenerReporteLiquidacionTicketHandler_TodosLosTurnos_RetornaListaVacia()
    {
        var parametros = new LiquidacionTicketParametros
        {
            TodosLosTurnos = true,
            FechaInicio = DateTime.Today.AddDays(-1),
            FechaFin = DateTime.Today
        };

        _repoMock.Setup(r => r.ObtenerLiquidacionTicketAsync(parametros, default))
            .ReturnsAsync(new List<LiquidacionTicketRow>().AsReadOnly());

        var handler = new ObtenerReporteLiquidacionTicketHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteLiquidacionTicketQuery(parametros));

        Assert.Empty(resultado.Filas);
        Assert.Equal("RepLiquidacionTicket.frx", resultado.NombrePlantilla);
    }

    [Fact]
    public async Task ObtenerReportePaloteoTicketHandler_OrdenCodigo_RetornaPlantillaCorrecta()
    {
        var parametros = new PaloteoTicketParametros
        {
            TodosTurnos = false,
            Turno = "T0001",
            Origen = OrigenPaloteoTicket.Produccion,
            OrdenarPorCodigoProducto = true
        };

        _repoMock.Setup(r => r.ObtenerPaloteoTicketAsync(parametros, default))
            .ReturnsAsync(new List<PaloteoTicketRow>
            {
                new() { TCodProducto = "P001", Producto = "Lomo", Cantidad = 2, Venta = 30 }
            }.AsReadOnly());

        var handler = new ObtenerReportePaloteoTicketHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReportePaloteoTicketQuery(parametros));

        Assert.Single(resultado.Filas);
        Assert.Equal("RepPaloteoTicket.frx", resultado.NombrePlantilla);
        Assert.Equal("Paloteo de Producción por Ticketera", resultado.TituloReporte);
        _repoMock.Verify(r => r.ObtenerPaloteoTicketAsync(parametros, default), Times.Once);
    }

    [Fact]
    public async Task ObtenerReportePaloteoTicketHandler_SinDatos_RetornaListaVacia()
    {
        var parametros = new PaloteoTicketParametros
        {
            TodosTurnos = true,
            FechaInicio = DateTime.Today.AddDays(-1),
            FechaFin = DateTime.Today,
            Origen = OrigenPaloteoTicket.PedidosFacturados
        };

        _repoMock.Setup(r => r.ObtenerPaloteoTicketAsync(parametros, default))
            .ReturnsAsync(new List<PaloteoTicketRow>().AsReadOnly());

        var handler = new ObtenerReportePaloteoTicketHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReportePaloteoTicketQuery(parametros));

        Assert.Empty(resultado.Filas);
        Assert.Equal("RepPaloteoTicket.frx", resultado.NombrePlantilla);
    }

    // ── BR-REP-017 DeliveryTicket ─────────────────────────────────────────────

    [Fact]
    public async Task ObtenerReporteDeliveryTicketHandler_ConResultados_RetornaNombrePlantillaYFilas()
    {
        var parametros = new DeliveryTicketParametros
        {
            TodosTurnos = false,
            Turno = "T001",
            TodasLasCajas = true,
            TodosLosMotorizados = true
        };

        _repoMock.Setup(r => r.ObtenerDeliveryTicketAsync(parametros, default))
            .ReturnsAsync(new List<DeliveryTicketRow>
            {
                new()
                {
                    TCaja = "001", TTipoPago = "01", TipoPago = "Efectivo",
                    TMotorizado = "M001", Motorizado = "Juan", TDocumento = "B001-0001",
                    FRegistro = DateTime.Today, NVenta = 50.00, NMonto = 50.00, NVuelto = 0,
                    TMoneda = "01", Mon = "S/.", TTurno = "T001", TUsuario = "USR1"
                }
            }.AsReadOnly());

        var handler = new ObtenerReporteDeliveryTicketHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteDeliveryTicketQuery(parametros));

        Assert.Single(resultado.Filas);
        Assert.Equal("RepDeliveryTicket.frx", resultado.NombrePlantilla);
        Assert.Equal("Cierre de Cajeros Delivery", resultado.TituloReporte);
        _repoMock.Verify(r => r.ObtenerDeliveryTicketAsync(parametros, default), Times.Once);
    }

    [Fact]
    public async Task ObtenerReporteDeliveryTicketHandler_SinResultados_RetornaListaVacia()
    {
        var parametros = new DeliveryTicketParametros
        {
            TodosTurnos = true,
            FechaInicio = DateTime.Today.AddDays(-1),
            FechaFin = DateTime.Today,
            TodasLasCajas = false,
            Caja = "002",
            TodosLosMotorizados = false,
            Motorizado = "M002"
        };

        _repoMock.Setup(r => r.ObtenerDeliveryTicketAsync(parametros, default))
            .ReturnsAsync(new List<DeliveryTicketRow>().AsReadOnly());

        var handler = new ObtenerReporteDeliveryTicketHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteDeliveryTicketQuery(parametros));

        Assert.Empty(resultado.Filas);
        Assert.Equal("RepDeliveryTicket.frx", resultado.NombrePlantilla);
    }

    // ── BR-REP-018 Reservas ───────────────────────────────────────────────────

    [Fact]
    public async Task ObtenerReporteReservasHandler_ConResultados_RetornaNombrePlantillaYFilas()
    {
        var parametros = new ReservaReporteParametros
        {
            FechaHoraInicio = new DateTime(2026, 8, 1, 8, 0, 0),
            FechaHoraFin    = new DateTime(2026, 8, 1, 23, 59, 59),
            EstadoGenerado  = true,
            EstadoAtendido  = false,
            EstadoAnulado   = false,
            Orden           = OrdenReserva.Fecha
        };

        _repoMock.Setup(r => r.ObtenerReservasReporteAsync(parametros, default))
            .ReturnsAsync(new List<ReservaReporteRow>
            {
                new()
                {
                    TReserva = "R001", Cliente = "Garcia Juan",
                    TTelefono = "999111222", NPax = 4,
                    FFecha = new DateTime(2026, 8, 1, 20, 0, 0),
                    TEstadoReserva = "01", EstadoReserva = "GENERADO"
                }
            }.AsReadOnly());

        var handler = new ObtenerReporteReservasHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteReservasQuery(parametros));

        Assert.Single(resultado.Filas);
        Assert.Equal("RepReservas.frx", resultado.NombrePlantilla);
        Assert.Equal("Reservas", resultado.TituloReporte);
        _repoMock.Verify(r => r.ObtenerReservasReporteAsync(parametros, default), Times.Once);
    }

    [Fact]
    public async Task ObtenerReporteReservasHandler_SinDatos_RetornaListaVacia()
    {
        var parametros = new ReservaReporteParametros
        {
            FechaHoraInicio = DateTime.Today,
            FechaHoraFin    = DateTime.Today.AddHours(23).AddMinutes(59),
            EstadoGenerado  = true,
            EstadoAtendido  = true,
            EstadoAnulado   = true,
            Orden           = OrdenReserva.Reserva
        };

        _repoMock.Setup(r => r.ObtenerReservasReporteAsync(parametros, default))
            .ReturnsAsync(new List<ReservaReporteRow>().AsReadOnly());

        var handler = new ObtenerReporteReservasHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteReservasQuery(parametros));

        Assert.Empty(resultado.Filas);
        Assert.Equal("RepReservas.frx", resultado.NombrePlantilla);
    }

    // ── BR-REP-019 Entregas ───────────────────────────────────────────────────

    [Fact]
    public async Task ObtenerReporteEntregaHandler_DetalladoFormato1_RetornaPlantillaCorrecta()
    {
        var parametros = new EntregaParametros
        {
            FechaHoraInicio = new DateTime(2026, 8, 1, 8, 0, 0),
            FechaHoraFin = new DateTime(2026, 8, 1, 23, 59, 59),
            CodigoCliente = "CLI001",
            Formato = FormatoReporteEntrega.DetalladoFormato1
        };

        _repoMock.Setup(r => r.ObtenerEntregasAsync(parametros, default))
            .ReturnsAsync(new List<EntregaRow>
            {
                new() { Pedido = "P001", Producto = "Lomo", Cantidad = 2, EstadoPedido = "ENTREGADO" }
            }.AsReadOnly());

        var handler = new ObtenerReporteEntregaHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteEntregaQuery(parametros));

        Assert.Single(resultado.Filas);
        Assert.Equal("RepEntregaFormato1.frx", resultado.NombrePlantilla);
        Assert.Equal("Reporte de Entregas", resultado.TituloReporte);
        _repoMock.Verify(r => r.ObtenerEntregasAsync(parametros, default), Times.Once);
    }

    [Fact]
    public async Task ObtenerReporteEntregaHandler_Resumido_RetornaPlantillaResumida()
    {
        var parametros = new EntregaParametros
        {
            FechaHoraInicio = DateTime.Today,
            FechaHoraFin = DateTime.Today.AddHours(23).AddMinutes(59),
            Formato = FormatoReporteEntrega.ResumidoPorProducto
        };

        _repoMock.Setup(r => r.ObtenerEntregasAsync(parametros, default))
            .ReturnsAsync(new List<EntregaRow>().AsReadOnly());

        var handler = new ObtenerReporteEntregaHandler(_repoMock.Object);
        var resultado = await handler.HandleAsync(new ObtenerReporteEntregaQuery(parametros));

        Assert.Empty(resultado.Filas);
        Assert.Equal("RepEntregaResumidoProd.frx", resultado.NombrePlantilla);
    }
}
