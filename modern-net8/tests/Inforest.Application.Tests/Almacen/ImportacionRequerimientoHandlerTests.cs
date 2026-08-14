using Inforest.Application.Almacen;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Almacen;
using Moq;

namespace Inforest.Application.Tests.Almacen;

/// <summary>
/// Tests de handlers POS-FUNC-014 — Importación de pedidos externos.
/// BR-IMPORT-001, BR-IMPORT-002, BR-IMPORT-003, BR-IMPORT-004.
/// </summary>
public class ImportacionRequerimientoHandlerTests
{
    // ── ObtenerRequerimientosPendientesHandler ───────────────────────────────

    [Fact]
    public async Task ObtenerPendientes_RangoValido_RetornaLista()
    {
        var rq1 = RequerimientoAlmacen.Crear("REQ-001", DateTime.Today, "Cocina", "Juan", "02");
        var repo = new Mock<IRequerimientoAlmacenRepository>();
        repo.Setup(r => r.ObtenerPendientesAsync(It.IsAny<DateTime>(), It.IsAny<DateTime>(), default))
            .ReturnsAsync(new List<RequerimientoAlmacen> { rq1 });

        var handler = new ObtenerRequerimientosPendientesHandler(repo.Object);
        var result = await handler.HandleAsync(
            new ObtenerRequerimientosPendientesQuery(DateTime.Today, DateTime.Today));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!);
    }

    [Fact]
    public async Task ObtenerPendientes_FechaFinAnteriorAInicio_RetornaError()
    {
        var repo = new Mock<IRequerimientoAlmacenRepository>();
        var handler = new ObtenerRequerimientosPendientesHandler(repo.Object);

        var result = await handler.HandleAsync(
            new ObtenerRequerimientosPendientesQuery(DateTime.Today, DateTime.Today.AddDays(-1)));

        Assert.False(result.EsExitoso);
        Assert.Equal("REQ_RANGO_FECHAS_INVALIDO", result.CodigoError);
    }

    // ── ObtenerDetalleRequerimientoHandler ───────────────────────────────────

    [Fact]
    public async Task ObtenerDetalle_RqValido_RetornaDetalle()
    {
        var d1 = DetalleRequerimientoAlmacen.Crear("ART-01", "Harina", 3.0, "PROD-01");
        var repo = new Mock<IRequerimientoAlmacenRepository>();
        repo.Setup(r => r.ObtenerDetalleAsync("REQ-001", "INFOREST", default))
            .ReturnsAsync(new List<DetalleRequerimientoAlmacen> { d1 });

        var handler = new ObtenerDetalleRequerimientoHandler(repo.Object);
        var result = await handler.HandleAsync(new ObtenerDetalleRequerimientoQuery("REQ-001"));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!);
        Assert.True(result.Valor![0].TieneEnlace);
    }

    [Fact]
    public async Task ObtenerDetalle_RqVacio_RetornaError()
    {
        var repo = new Mock<IRequerimientoAlmacenRepository>();
        var handler = new ObtenerDetalleRequerimientoHandler(repo.Object);

        var result = await handler.HandleAsync(new ObtenerDetalleRequerimientoQuery(""));

        Assert.False(result.EsExitoso);
        Assert.Equal("REQ_CODIGO_REQUERIDO", result.CodigoError);
    }

    // ── ImportarRequerimientoHandler ─────────────────────────────────────────

    [Fact]
    public async Task Importar_RequerimientoYaImportado_RetornaError_BR_IMPORT_002()
    {
        // BR-IMPORT-002: Si ya tiene pedido, debe fallar
        var repo = new Mock<IRequerimientoAlmacenRepository>();
        repo.Setup(r => r.ObtenerPedidoAsociadoAsync("REQ-001", default))
            .ReturnsAsync("PED-0001");

        var gateway = new Mock<IImportacionPedidoGateway>();
        var handler = new ImportarRequerimientoHandler(repo.Object, gateway.Object);

        var cmd = CrearComando("REQ-001");
        var result = await handler.HandleAsync(cmd);

        Assert.False(result.EsExitoso);
        Assert.Equal("REQ_YA_IMPORTADO", result.CodigoError);
        gateway.Verify(g => g.CrearPedidoDesdeRequerimientoAsync(It.IsAny<ImportacionPedidoContexto>(), default), Times.Never);
    }

    [Fact]
    public async Task Importar_SinProductos_RetornaError()
    {
        var repo = new Mock<IRequerimientoAlmacenRepository>();
        repo.Setup(r => r.ObtenerPedidoAsociadoAsync("REQ-001", default))
            .ReturnsAsync((string?)null);
        repo.Setup(r => r.ObtenerDetalleAsync("REQ-001", "INFOREST", default))
            .ReturnsAsync(new List<DetalleRequerimientoAlmacen>());

        var gateway = new Mock<IImportacionPedidoGateway>();
        var handler = new ImportarRequerimientoHandler(repo.Object, gateway.Object);

        var result = await handler.HandleAsync(CrearComando("REQ-001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("REQ_SIN_PRODUCTOS", result.CodigoError);
    }

    [Fact]
    public async Task Importar_ProductoSinEnlace_RetornaError_BR_IMPORT_003()
    {
        // BR-IMPORT-003: Si algún artículo no tiene enlace, debe fallar antes de crear el pedido
        var detalleSinEnlace = DetalleRequerimientoAlmacen.Crear("ART-99", "Sin enlace", 2.0);

        var repo = new Mock<IRequerimientoAlmacenRepository>();
        repo.Setup(r => r.ObtenerPedidoAsociadoAsync("REQ-001", default))
            .ReturnsAsync((string?)null);
        repo.Setup(r => r.ObtenerDetalleAsync("REQ-001", "INFOREST", default))
            .ReturnsAsync(new List<DetalleRequerimientoAlmacen> { detalleSinEnlace });

        var gateway = new Mock<IImportacionPedidoGateway>();
        var handler = new ImportarRequerimientoHandler(repo.Object, gateway.Object);

        var result = await handler.HandleAsync(CrearComando("REQ-001"));

        Assert.False(result.EsExitoso);
        Assert.Equal("REQ_PRODUCTO_SIN_ENLACE", result.CodigoError);
        gateway.Verify(g => g.CrearPedidoDesdeRequerimientoAsync(It.IsAny<ImportacionPedidoContexto>(), default), Times.Never);
    }

    [Fact]
    public async Task Importar_Exitoso_MarcaRequerimiento_BR_IMPORT_004()
    {
        // BR-IMPORT-004: Tras importar exitosamente, debe marcarse el requerimiento
        var detalle = DetalleRequerimientoAlmacen.Crear("ART-01", "Harina", 3.0, "PROD-01");

        var repo = new Mock<IRequerimientoAlmacenRepository>();
        repo.Setup(r => r.ObtenerPedidoAsociadoAsync("REQ-001", default))
            .ReturnsAsync((string?)null);
        repo.Setup(r => r.ObtenerDetalleAsync("REQ-001", "INFOREST", default))
            .ReturnsAsync(new List<DetalleRequerimientoAlmacen> { detalle });

        var gateway = new Mock<IImportacionPedidoGateway>();
        gateway.Setup(g => g.CrearPedidoDesdeRequerimientoAsync(
                It.IsAny<ImportacionPedidoContexto>(), default))
            .ReturnsAsync(Domain.Common.Result.Ok(new ImportarRequerimientoResult("PED-0099", 1)));

        var handler = new ImportarRequerimientoHandler(repo.Object, gateway.Object);
        var result = await handler.HandleAsync(CrearComando("REQ-001"));

        Assert.True(result.EsExitoso);
        Assert.Equal("PED-0099", result.Valor!.CodigoPedido);
        Assert.Equal(1, result.Valor!.ProductosImportados);

        // BR-IMPORT-004: Debe marcarse como importado
        repo.Verify(r => r.MarcarImportadoAsync("REQ-001", "PED-0099", default), Times.Once);
    }

    [Fact]
    public async Task Importar_GatewayFalla_NoMarcaRequerimiento()
    {
        var detalle = DetalleRequerimientoAlmacen.Crear("ART-01", "Harina", 3.0, "PROD-01");

        var repo = new Mock<IRequerimientoAlmacenRepository>();
        repo.Setup(r => r.ObtenerPedidoAsociadoAsync("REQ-001", default))
            .ReturnsAsync((string?)null);
        repo.Setup(r => r.ObtenerDetalleAsync("REQ-001", "INFOREST", default))
            .ReturnsAsync(new List<DetalleRequerimientoAlmacen> { detalle });

        var gateway = new Mock<IImportacionPedidoGateway>();
        gateway.Setup(g => g.CrearPedidoDesdeRequerimientoAsync(
                It.IsAny<ImportacionPedidoContexto>(), default))
            .ReturnsAsync(Domain.Common.Result.Fail<ImportarRequerimientoResult>(
                "Producto no encontrado", "IMPORT_REQ_PRODUCTO_NO_ENCONTRADO"));

        var handler = new ImportarRequerimientoHandler(repo.Object, gateway.Object);
        var result = await handler.HandleAsync(CrearComando("REQ-001"));

        Assert.False(result.EsExitoso);
        repo.Verify(r => r.MarcarImportadoAsync(It.IsAny<string>(), It.IsAny<string>(), default), Times.Never);
    }

    // ── helpers ──────────────────────────────────────────────────────────────

    private static ImportarRequerimientoCommand CrearComando(string rq)
        => new(
            Rq: rq,
            CodigoCaja: "001",
            CodigoTurno: "T001",
            CodigoSalon: "01",
            CodigoUsuario: "ADMIN",
            TipoPedido: "01",
            FechaDiaContable: DateTime.Today,
            BaseDatosInforest: "INFOREST");
}
