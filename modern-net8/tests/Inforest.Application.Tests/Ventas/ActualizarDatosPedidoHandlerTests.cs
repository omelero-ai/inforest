using Inforest.Application.Ventas;
using Moq;

namespace Inforest.Application.Tests.Ventas;

/// <summary>
/// Tests de los handlers de actualización de datos del pedido.
/// BR-ACTPED-001/002/003.
/// Legacy: FrmActualizarPedidos (frmUpdateDatosPedido.frm).
/// </summary>
public class ActualizarDatosPedidoHandlerTests
{
    private static DatosPedidoActualizar BuildDatos(string cortesia = "01", string descCortesia = "Invitado", string canal = "01", string descCanal = "Salón")
        => new("PED001", cortesia, descCortesia, canal, descCanal);

    // ── ActualizarCortesiaPedidoHandler ──────────────────────────────────────

    [Fact]
    public async Task ActualizarCortesia_SinCortesiaPrevia_RetornaError_ACTPED002()
    {
        var repo = new Mock<IActualizarDatosPedidoRepository>();
        var handler = new ActualizarCortesiaPedidoHandler(repo.Object);

        var result = await handler.HandleAsync(new ActualizarCortesiaPedidoCommand(
            "DOC001", "", "02", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("ACTPED_SIN_CORTESIA", result.CodigoError);
    }

    [Fact]
    public async Task ActualizarCortesia_SinNuevaCortesia_RetornaError()
    {
        var repo = new Mock<IActualizarDatosPedidoRepository>();
        var handler = new ActualizarCortesiaPedidoHandler(repo.Object);

        var result = await handler.HandleAsync(new ActualizarCortesiaPedidoCommand(
            "DOC001", "01", "", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("ACTPED_CORTESIA_REQUERIDA", result.CodigoError);
    }

    [Fact]
    public async Task ActualizarCortesia_Exitoso_LlamaRepo()
    {
        var repo = new Mock<IActualizarDatosPedidoRepository>();
        repo.Setup(r => r.ActualizarCortesiaAsync(
                "DOC001", "02", "01", "USR01", It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var handler = new ActualizarCortesiaPedidoHandler(repo.Object);
        var result = await handler.HandleAsync(new ActualizarCortesiaPedidoCommand(
            "DOC001", "01", "02", "USR01"));

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.ActualizarCortesiaAsync("DOC001", "02", "01", "USR01", It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ActualizarCortesia_ExcepcionRepo_RetornaError()
    {
        var repo = new Mock<IActualizarDatosPedidoRepository>();
        repo.Setup(r => r.ActualizarCortesiaAsync(
                It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("DB error"));

        var handler = new ActualizarCortesiaPedidoHandler(repo.Object);
        var result = await handler.HandleAsync(new ActualizarCortesiaPedidoCommand(
            "DOC001", "01", "02", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("ACTPED_CORTESIA_ERROR", result.CodigoError);
    }

    // ── ActualizarCanalVentaPedidoHandler ────────────────────────────────────

    [Fact]
    public async Task ActualizarCanal_SinCanalNuevo_RetornaError()
    {
        var repo = new Mock<IActualizarDatosPedidoRepository>();
        var handler = new ActualizarCanalVentaPedidoHandler(repo.Object);

        var result = await handler.HandleAsync(new ActualizarCanalVentaPedidoCommand(
            "DOC001", "01", "", "", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("ACTPED_CANAL_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task ActualizarCanal_DeliverySinCliente_RetornaError_ACTPED003()
    {
        var repo = new Mock<IActualizarDatosPedidoRepository>();
        var handler = new ActualizarCanalVentaPedidoHandler(repo.Object);

        // Canal 02 = Delivery, sin cliente → error
        var result = await handler.HandleAsync(new ActualizarCanalVentaPedidoCommand(
            "DOC001", "01", "02", "", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("ACTPED_CANAL_DELIVERY_SIN_CLIENTE", result.CodigoError);
    }

    [Fact]
    public async Task ActualizarCanal_Exitoso_LlamaRepo()
    {
        var repo = new Mock<IActualizarDatosPedidoRepository>();
        repo.Setup(r => r.ActualizarCanalVentaAsync(
                "DOC001", "03", "", "01", "USR01", It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var handler = new ActualizarCanalVentaPedidoHandler(repo.Object);
        var result = await handler.HandleAsync(new ActualizarCanalVentaPedidoCommand(
            "DOC001", "01", "03", "", "USR01"));

        Assert.True(result.EsExitoso);
    }

    [Fact]
    public async Task ActualizarCanal_DeliveryConCliente_Exitoso()
    {
        var repo = new Mock<IActualizarDatosPedidoRepository>();
        repo.Setup(r => r.ActualizarCanalVentaAsync(
                "DOC001", "02", "CLI001", "01", "USR01", It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var handler = new ActualizarCanalVentaPedidoHandler(repo.Object);
        var result = await handler.HandleAsync(new ActualizarCanalVentaPedidoCommand(
            "DOC001", "01", "02", "CLI001", "USR01"));

        Assert.True(result.EsExitoso);
    }
}
