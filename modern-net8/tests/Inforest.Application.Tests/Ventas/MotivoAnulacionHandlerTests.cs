using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Application.Tests.Ventas;

/// <summary>
/// Tests para ObtenerMotivosAnulacionHandler.
/// Legacy: frmListaMotivos.frm / vMotivoAnulacion.
/// BR-MOTIVO-001: solo activos. BR-MOTIVO-002: falla si vacío. BR-MOTIVO-003: retorna motivo.
/// </summary>
public sealed class MotivoAnulacionHandlerTests
{
    private static ObtenerMotivosAnulacionHandler BuildHandler(
        IReadOnlyList<MotivoAnulacion> motivos)
    {
        var repo = new StubMotivoAnulacionRepository(motivos);
        return new ObtenerMotivosAnulacionHandler(repo);
    }

    // BR-MOTIVO-001: devuelve lista de motivos activos
    [Fact]
    public async Task HandleAsync_ConMotivosActivos_DevueltaLista()
    {
        var motivos = new List<MotivoAnulacion>
        {
            new("001", "Error de Producto", "Error", true),
            new("002", "Cambio de Pedido", "Cambio", true),
            new("003", "Solicitud del Cliente", "Cliente", true)
        };
        var handler = BuildHandler(motivos);

        var result = await handler.HandleAsync(new ObtenerMotivosAnulacionQuery());

        Assert.True(result.EsExitoso);
        Assert.Equal(3, result.Valor!.Motivos.Count);
    }

    // BR-MOTIVO-001: los códigos y descripciones se preservan
    [Fact]
    public async Task HandleAsync_VerificaCamposMotivo()
    {
        var motivos = new List<MotivoAnulacion>
        {
            new("001", "Error de Producto", "Error", true, 0)
        };
        var handler = BuildHandler(motivos);

        var result = await handler.HandleAsync(new ObtenerMotivosAnulacionQuery());

        Assert.True(result.EsExitoso);
        var motivo = result.Valor!.Motivos[0];
        Assert.Equal("001", motivo.Codigo);
        Assert.Equal("Error de Producto", motivo.Descripcion);
        Assert.Equal("Error", motivo.Resumido);
        Assert.True(motivo.Activo);
    }

    // BR-MOTIVO-002: sin motivos configurados retorna falla
    [Fact]
    public async Task HandleAsync_SinMotivos_RetornaFalla()
    {
        var handler = BuildHandler([]);

        var result = await handler.HandleAsync(new ObtenerMotivosAnulacionQuery());

        Assert.False(result.EsExitoso);
        Assert.Equal("MOTIVO_VACIO", result.CodigoError);
        Assert.NotNull(result.MensajeError);
    }

    // BR-MOTIVO-002: mensaje de error legible para el operador
    [Fact]
    public async Task HandleAsync_SinMotivos_MensajeEnEspanol()
    {
        var handler = BuildHandler([]);

        var result = await handler.HandleAsync(new ObtenerMotivosAnulacionQuery());

        Assert.False(result.EsExitoso);
        Assert.Contains("Motivos", result.MensajeError, StringComparison.OrdinalIgnoreCase);
    }

    // BR-MOTIVO-003: con un solo motivo también retorna éxito
    [Fact]
    public async Task HandleAsync_UnMotivo_DevuelveExito()
    {
        var motivos = new List<MotivoAnulacion> { new("001", "Error", "Err", true) };
        var handler = BuildHandler(motivos);

        var result = await handler.HandleAsync(new ObtenerMotivosAnulacionQuery());

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!.Motivos);
    }

    // Stub de repositorio
    private sealed class StubMotivoAnulacionRepository(IReadOnlyList<MotivoAnulacion> data)
        : IMotivoAnulacionRepository
    {
        public Task<IReadOnlyList<MotivoAnulacion>> ObtenerActivosAsync(CancellationToken ct = default)
            => Task.FromResult(data);
    }
}
