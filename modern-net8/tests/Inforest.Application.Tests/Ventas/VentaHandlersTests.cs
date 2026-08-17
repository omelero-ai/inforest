using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;
using Moq;

namespace Inforest.Application.Tests.Ventas;

public class EmitirDocumentoHandlerTests
{
    private static Pedido BuildPedidoConDetalle(string codigoPedido = "PED001", string codigoTurno = "TUR001", string codigoCaja = "CAJ01")
    {
        var pedido = Pedido.Crear(codigoPedido, "M01", codigoTurno, "USR01", CanalVenta.Local);
        pedido.AgregarDetalle(DetallePedido.Crear(codigoPedido, "001", "PROD01", 2m, 10m, 10m));
        return pedido;
    }

    [Fact]
    public async Task EmitirDocumento_PedidoValido_RetornaDocumento()
    {
        var pedido = BuildPedidoConDetalle();

        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Documento?)null);
        docRepo.Setup(r => r.InsertarAsync(It.IsAny<Documento>(), It.IsAny<IReadOnlyList<DetalleDocumento>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var sessionService = new Mock<ISessionService>();
        sessionService.Setup(s => s.SesionActual)
            .Returns(BuildSesion("CAJ01"));
        var clienteRepo = new Mock<IClienteRepository>();

        var handler = new EmitirDocumentoHandler(pedidoRepo.Object, docRepo.Object, sessionService.Object, clienteRepo.Object);
        var result = await handler.HandleAsync(new EmitirDocumentoCommand("PED001", "01", null, 0m, 0m));

        Assert.True(result.EsExitoso);
        Assert.NotNull(result.Valor);
        Assert.Equal("01", result.Valor!.TipoDocumento);
        docRepo.Verify(r => r.InsertarAsync(It.IsAny<Documento>(), It.IsAny<IReadOnlyList<DetalleDocumento>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task EmitirDocumento_PedidoSinDetalles_RetornaError()
    {
        var pedido = Pedido.Crear("PED002", "M01", "TUR001", "USR01", CanalVenta.Local);

        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED002", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var docRepo = new Mock<IDocumentoRepository>();
        var sessionService = new Mock<ISessionService>();
        sessionService.Setup(s => s.SesionActual)
            .Returns(BuildSesion("CAJ01"));
        var clienteRepo = new Mock<IClienteRepository>();

        var handler = new EmitirDocumentoHandler(pedidoRepo.Object, docRepo.Object, sessionService.Object, clienteRepo.Object);
        var result = await handler.HandleAsync(new EmitirDocumentoCommand("PED002", "01", null, 0m, 0m));

        Assert.False(result.EsExitoso);
        Assert.Equal("VENTA_PEDIDO_SIN_DETALLES", result.CodigoError);
    }

    [Fact]
    public async Task EmitirDocumento_PedidoNoExiste_RetornaError()
    {
        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED_NOPE", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Pedido?)null);

        var docRepo = new Mock<IDocumentoRepository>();
        var sessionService = new Mock<ISessionService>();
        sessionService.Setup(s => s.SesionActual).Returns(BuildSesion("CAJ01"));
        var clienteRepo = new Mock<IClienteRepository>();

        var handler = new EmitirDocumentoHandler(pedidoRepo.Object, docRepo.Object, sessionService.Object, clienteRepo.Object);
        var result = await handler.HandleAsync(new EmitirDocumentoCommand("PED_NOPE", "01", null, 0m, 0m));

        Assert.False(result.EsExitoso);
        Assert.Equal("VENTA_PEDIDO_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task EmitirDocumento_SinSesion_RetornaError()
    {
        var pedido = BuildPedidoConDetalle();

        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Documento?)null);

        var sessionService = new Mock<ISessionService>();
        sessionService.Setup(s => s.SesionActual).Returns((Domain.Entities.Seguridad.SesionOperativa?)null);
        var clienteRepo = new Mock<IClienteRepository>();

        var handler = new EmitirDocumentoHandler(pedidoRepo.Object, docRepo.Object, sessionService.Object, clienteRepo.Object);
        var result = await handler.HandleAsync(new EmitirDocumentoCommand("PED001", "01", null, 0m, 0m));

        Assert.False(result.EsExitoso);
        Assert.Equal("VENTA_CAJA_NO_DISPONIBLE", result.CodigoError);
    }

    [Fact]
    public async Task EmitirDocumento_DocumentoDuplicado_RetornaError()
    {
        var pedido = BuildPedidoConDetalle();
        var docExistente = Documento.Emitir("DOC_EXIST", "01", null, 10m, 0m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");

        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(docExistente);

        var sessionService = new Mock<ISessionService>();
        sessionService.Setup(s => s.SesionActual).Returns(BuildSesion("CAJ01"));
        var clienteRepo = new Mock<IClienteRepository>();

        var handler = new EmitirDocumentoHandler(pedidoRepo.Object, docRepo.Object, sessionService.Object, clienteRepo.Object);
        var result = await handler.HandleAsync(new EmitirDocumentoCommand("PED001", "01", null, 0m, 0m));

        Assert.False(result.EsExitoso);
        Assert.Equal("VENTA_DOCUMENTO_DUPLICADO", result.CodigoError);
    }

    [Fact]
    public async Task EmitirDocumento_ClienteIncompatible_RetornaError()
    {
        var pedido = BuildPedidoConDetalle();

        var pedidoRepo = new Mock<IPedidoRepository>();
        pedidoRepo.Setup(r => r.ObtenerPorCodigoAsync("PED001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(pedido);

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Documento?)null);

        var sessionService = new Mock<ISessionService>();
        sessionService.Setup(s => s.SesionActual).Returns(BuildSesion("CAJ01"));

        var clienteRepo = new Mock<IClienteRepository>();
        clienteRepo.Setup(r => r.ValidarCompatibilidadDocumentoAsync("03", "CLI001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(false);

        var handler = new EmitirDocumentoHandler(pedidoRepo.Object, docRepo.Object, sessionService.Object, clienteRepo.Object);
        var result = await handler.HandleAsync(new EmitirDocumentoCommand("PED001", "03", "CLI001", 0m, 0m));

        Assert.False(result.EsExitoso);
        Assert.Equal("VENTA_CLIENTE_TIPO_DOCUMENTO_INVALIDO", result.CodigoError);
        docRepo.Verify(r => r.InsertarAsync(It.IsAny<Documento>(), It.IsAny<IReadOnlyList<DetalleDocumento>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    private static Domain.Entities.Seguridad.SesionOperativa BuildSesion(string codigoCaja)
        => Domain.Entities.Seguridad.SesionOperativa.Iniciar(
            "USR01", "Usuario Prueba", "GRP01", "INFOREST",
            codigoCaja, "TERM01", "INFOREST", 1, DateTime.Now, []);
}

public class AnularDocumentoHandlerTests
{
    [Fact]
    public async Task AnularDocumento_Exitoso()
    {
        var documento = Documento.Emitir("DOC001", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);
        docRepo.Setup(r => r.AnularAsync("DOC001", "USR01", "Test", It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = new AnularDocumentoHandler(docRepo.Object);
        var result = await handler.HandleAsync(new AnularDocumentoCommand("DOC001", "USR01", "Test"));

        Assert.True(result.EsExitoso);
        docRepo.Verify(r => r.AnularAsync("DOC001", "USR01", "Test", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task AnularDocumento_NoExiste_RetornaError()
    {
        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC_NOPE", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Documento?)null);

        var handler = new AnularDocumentoHandler(docRepo.Object);
        var result = await handler.HandleAsync(new AnularDocumentoCommand("DOC_NOPE", "USR01", "Test"));

        Assert.False(result.EsExitoso);
        Assert.Equal("VENTA_DOCUMENTO_NO_ENCONTRADO", result.CodigoError);
    }

    [Fact]
    public async Task AnularDocumento_YaAnulado_LanzaExcepcionDominio()
    {
        var documento = Documento.Emitir("DOC001", "01", null, 100m, 18m, 0m, 0m, 0m, 0m, "TUR001", "CAJ01");
        documento.Anular("USR01", "Primera anulación");

        var docRepo = new Mock<IDocumentoRepository>();
        docRepo.Setup(r => r.ObtenerPorCodigoAsync("DOC001", It.IsAny<CancellationToken>()))
            .ReturnsAsync(documento);

        var handler = new AnularDocumentoHandler(docRepo.Object);
        var result = await handler.HandleAsync(new AnularDocumentoCommand("DOC001", "USR02", "Segunda anulación"));

        // DomainException capturada → resultado fallido
        Assert.False(result.EsExitoso);
        Assert.Equal("DOCUMENTO_YA_ANULADO", result.CodigoError);
    }
}
