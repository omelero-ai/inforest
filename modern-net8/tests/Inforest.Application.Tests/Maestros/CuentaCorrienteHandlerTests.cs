using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;
using Moq;

namespace Inforest.Application.Tests.Maestros;

/// <summary>
/// Tests de handlers POS-FUNC-010 — Cliente y cuentas corrientes.
/// BR-CLIENTE-001..004, BR-CTACTE-001..003.
/// </summary>
public class CuentaCorrienteHandlerTests
{
    // ── ObtenerCuentasCorrientesHandler ──────────────────────────────────────

    [Fact]
    public async Task ObtenerCuentasCorrientes_RetornaLista()
    {
        var ctacte1 = CuentaCorriente.Crear("DLV001", "Empresa A", 1000m, 200m, "user");
        var ctacte2 = CuentaCorriente.Crear("DLV002", "Empresa B", 500m, 0m, "user");

        var repo = new Mock<ICuentaCorrienteRepository>();
        repo.Setup(r => r.ObtenerTodosAsync(default))
            .ReturnsAsync(new List<CuentaCorriente> { ctacte1, ctacte2 });

        var handler = new ObtenerCuentasCorrientesHandler(repo.Object);
        var result  = await handler.HandleAsync(new ObtenerCuentasCorrientesQuery());

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count);
    }

    // ── ObtenerCuentaCorrientePorCodigoHandler ────────────────────────────────

    [Fact]
    public async Task ObtenerPorCodigo_Existente_RetornaEntidad()
    {
        var ctacte = CuentaCorriente.Crear("DLV001", "Empresa A", 1000m, 200m, "user");
        var repo = new Mock<ICuentaCorrienteRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("DLV001", default))
            .ReturnsAsync(ctacte);

        var handler = new ObtenerCuentaCorrientePorCodigoHandler(repo.Object);
        var result  = await handler.HandleAsync(new ObtenerCuentaCorrientePorCodigoQuery("DLV001"));

        Assert.True(result.EsExitoso);
        Assert.Equal("DLV001", result.Valor!.CodigoCtaCte);
    }

    [Fact]
    public async Task ObtenerPorCodigo_Inexistente_RetornaFallo()
    {
        var repo = new Mock<ICuentaCorrienteRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("NOEXISTE", default))
            .ReturnsAsync((CuentaCorriente?)null);

        var handler = new ObtenerCuentaCorrientePorCodigoHandler(repo.Object);
        var result  = await handler.HandleAsync(new ObtenerCuentaCorrientePorCodigoQuery("NOEXISTE"));

        Assert.False(result.EsExitoso);
        Assert.Equal("CTACTE_NO_ENCONTRADA", result.CodigoError);
    }

    // ── ObtenerDocumentosPendientesCobroHandler ───────────────────────────────

    [Fact]
    public async Task ObtenerDocumentosPendientes_FechaValida_RetornaLista()
    {
        var doc1 = new DocumentoPendienteCobro("DOC001", "CAJ1", "CLI01", "Cliente A", null, null, 150m, DateTime.Today, "TURNO1", "user", "SIN VENCER");
        var repo = new Mock<ICuentaCorrienteRepository>();
        repo.Setup(r => r.ObtenerDocumentosPendientesCobroAsync(
                It.IsAny<DateTime>(), It.IsAny<DateTime>(), default))
            .ReturnsAsync(new List<DocumentoPendienteCobro> { doc1 });

        var handler = new ObtenerDocumentosPendientesCobroHandler(repo.Object);
        var result  = await handler.HandleAsync(
            new ObtenerDocumentosPendientesCobroQuery(DateTime.Today.AddDays(-7), DateTime.Today));

        Assert.True(result.EsExitoso);
        Assert.Single(result.Valor!);
        Assert.Equal("DOC001", result.Valor![0].Documento);
    }

    [Fact]
    public async Task ObtenerDocumentosPendientes_FechaFinAnteriorFechaInicio_RetornaFallo()
    {
        var repo = new Mock<ICuentaCorrienteRepository>();
        var handler = new ObtenerDocumentosPendientesCobroHandler(repo.Object);
        var result  = await handler.HandleAsync(
            new ObtenerDocumentosPendientesCobroQuery(DateTime.Today, DateTime.Today.AddDays(-1)));

        Assert.False(result.EsExitoso);
        Assert.Equal("CTACTE_FECHA_INVALIDA", result.CodigoError);
    }

    // ── RegistrarNuevoClientePosHandler ──────────────────────────────────────

    [Fact]
    public async Task RegistrarNuevoCliente_CodigoNuevo_Insertado()
    {
        var cliente = Cliente.Crear("CLI99", "Nuevo Cliente", "12345678", "Av. Test 1", "tester", tipoIdentidad: "DNI");
        var repo = new Mock<IClienteRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("CLI99", default))
            .ReturnsAsync((Cliente?)null);
        repo.Setup(r => r.InsertarAsync(It.IsAny<Cliente>(), default))
            .ReturnsAsync(true);

        var handler = new RegistrarNuevoClientePosHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarNuevoClientePosCommand(cliente));

        Assert.True(result.EsExitoso);
        repo.Verify(r => r.InsertarAsync(It.IsAny<Cliente>(), default), Times.Once);
    }

    [Fact]
    public async Task RegistrarNuevoCliente_CodigoExistente_RetornaFallo()
    {
        var clienteExistente = Cliente.Crear("CLI01", "Existe", "12345678", "Av. Exist 1", "tester", tipoIdentidad: "DNI");
        var nuevoCliente     = Cliente.Crear("CLI01", "Nuevo",  "87654321", "Av. New 1",  "tester", tipoIdentidad: "DNI");

        var repo = new Mock<IClienteRepository>();
        repo.Setup(r => r.ObtenerPorCodigoAsync("CLI01", default))
            .ReturnsAsync(clienteExistente);

        var handler = new RegistrarNuevoClientePosHandler(repo.Object);
        var result  = await handler.HandleAsync(new RegistrarNuevoClientePosCommand(nuevoCliente));

        Assert.False(result.EsExitoso);
        Assert.Equal("CLIENTE_YA_EXISTE", result.CodigoError);
        repo.Verify(r => r.InsertarAsync(It.IsAny<Cliente>(), default), Times.Never);
    }
}
