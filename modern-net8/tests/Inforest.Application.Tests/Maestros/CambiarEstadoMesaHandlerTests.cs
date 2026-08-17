using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;
using Moq;

namespace Inforest.Application.Tests.Maestros;

/// <summary>
/// Tests del handler de cambio de estado de mesa.
/// Legacy: frmMesaConsulta.frm — UPDATE TMESA SET tEstadoMesa='XX'.
/// BR-MESACONSULTA-001.
/// </summary>
public class CambiarEstadoMesaHandlerTests
{
    private static Mesa CrearMesa(string codigo, EstadoMesa estado)
    {
        var mesa = Mesa.Crear(codigo, "Mesa " + codigo, null, "S01", false, null, null, 4, "USR");
        if (estado == EstadoMesa.Ocupada)  mesa.Ocupar();
        if (estado == EstadoMesa.Reservada) mesa.Reservar();
        return mesa;
    }

    [Fact]
    public async Task Handle_MesaNoExiste_RetornaFallo()
    {
        var repo = new Mock<IMesaRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("M99", It.IsAny<CancellationToken>()))
            .ReturnsAsync((Mesa?)null);

        var handler = new CambiarEstadoMesaHandler(repo.Object);
        var r = await handler.HandleAsync(new CambiarEstadoMesaCommand("M99", EstadoMesa.Sucia));

        Assert.False(r.EsExitoso);
        Assert.Equal("MESA_NO_ENCONTRADA", r.CodigoError);
    }

    [Fact]
    public async Task Handle_MesaOcupada_NoPermiteCambiarEstado()
    {
        // BR-MESACONSULTA-001: no se puede cambiar estado de una mesa ocupada
        var repo = new Mock<IMesaRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("M01", It.IsAny<CancellationToken>()))
            .ReturnsAsync(CrearMesa("M01", EstadoMesa.Ocupada));

        var handler = new CambiarEstadoMesaHandler(repo.Object);
        var r = await handler.HandleAsync(new CambiarEstadoMesaCommand("M01", EstadoMesa.Libre));

        Assert.False(r.EsExitoso);
        Assert.Equal("MESA_OCUPADA", r.CodigoError);
        repo.Verify(x => x.CambiarEstadoAsync(It.IsAny<string>(), It.IsAny<EstadoMesa>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task Handle_MesaLibre_CambiaEstadoAReservada()
    {
        var repo = new Mock<IMesaRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("M02", It.IsAny<CancellationToken>()))
            .ReturnsAsync(CrearMesa("M02", EstadoMesa.Libre));
        repo.Setup(x => x.CambiarEstadoAsync("M02", EstadoMesa.Reservada, It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = new CambiarEstadoMesaHandler(repo.Object);
        var r = await handler.HandleAsync(new CambiarEstadoMesaCommand("M02", EstadoMesa.Reservada));

        Assert.True(r.EsExitoso);
        repo.Verify(x => x.CambiarEstadoAsync("M02", EstadoMesa.Reservada, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Handle_MesaSucia_CambiaEstadoALibre()
    {
        var mesa = Mesa.Crear("M03", "Mesa 03", null, "S01", false, null, null, 4, "USR",
            estado: EstadoMesa.Sucia);
        var repo = new Mock<IMesaRepository>();
        repo.Setup(x => x.ObtenerPorCodigoAsync("M03", It.IsAny<CancellationToken>()))
            .ReturnsAsync(mesa);
        repo.Setup(x => x.CambiarEstadoAsync("M03", EstadoMesa.Libre, It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var handler = new CambiarEstadoMesaHandler(repo.Object);
        var r = await handler.HandleAsync(new CambiarEstadoMesaCommand("M03", EstadoMesa.Libre));

        Assert.True(r.EsExitoso);
    }
}
