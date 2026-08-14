using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Ventas;

/// <summary>
/// Tests de la entidad NotaCredito.
/// BR-NC-001: Total > 0.
/// BR-NC-002: Impuestos >= 0.
/// BR-NC-004: Observacion requerida.
/// BR-NC-006: No reanular NC anulada.
/// </summary>
public class NotaCreditoTests
{
    // BR-NC-001
    [Fact]
    public void Emitir_TotalCero_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            NotaCredito.Emitir("NC001", DateTime.Today, "DOC001", 0m, 0m, 0m, 0m, "Motivo prueba"));

        Assert.Equal("NC_TOTAL_INVALIDO", ex.Code);
    }

    // BR-NC-001
    [Fact]
    public void Emitir_TotalPositivo_EstadoInicial_PE()
    {
        var nc = NotaCredito.Emitir("NC001", DateTime.Today, "DOC001", 100m, 18m, 0m, 0m, "Motivo prueba");

        Assert.Equal("PE", nc.Estado);
        Assert.Equal(118m, nc.Total);
    }

    // BR-NC-002
    [Fact]
    public void Emitir_ImpuestoNegativo_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            NotaCredito.Emitir("NC002", DateTime.Today, "DOC001", 100m, -1m, 0m, 0m, "Motivo prueba"));

        Assert.Equal("NC_IMPUESTO_NEGATIVO", ex.Code);
    }

    // BR-NC-004
    [Fact]
    public void Emitir_ObservacionVacia_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            NotaCredito.Emitir("NC003", DateTime.Today, "DOC001", 100m, 18m, 0m, 0m, ""));

        Assert.Equal("NC_OBSERVACION_REQUERIDA", ex.Code);
    }

    [Fact]
    public void Emitir_CodigoBlancos_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            NotaCredito.Emitir("  ", DateTime.Today, "DOC001", 100m, 18m, 0m, 0m, "Motivo"));

        Assert.Equal("NC_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Emitir_DocumentoVacio_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            NotaCredito.Emitir("NC004", DateTime.Today, "", 100m, 18m, 0m, 0m, "Motivo"));

        Assert.Equal("NC_DOCUMENTO_REQUERIDO", ex.Code);
    }

    // BR-NC-006
    [Fact]
    public void Anular_NCYaAnulada_LanzaExcepcion()
    {
        var nc = NotaCredito.Emitir("NC005", DateTime.Today, "DOC001", 50m, 9m, 0m, 0m, "Motivo");
        nc.Anular("ADMIN");

        var ex = Assert.Throws<DomainException>(() => nc.Anular("ADMIN"));
        Assert.Equal("NC_YA_ANULADA", ex.Code);
    }

    [Fact]
    public void Anular_UsuarioVacio_LanzaExcepcion()
    {
        var nc = NotaCredito.Emitir("NC006", DateTime.Today, "DOC001", 50m, 9m, 0m, 0m, "Motivo");

        var ex = Assert.Throws<DomainException>(() => nc.Anular(""));
        Assert.Equal("NC_USUARIO_ANULACION_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Anular_Exitosa_EstadoCambiaAN()
    {
        var nc = NotaCredito.Emitir("NC007", DateTime.Today, "DOC001", 50m, 9m, 0m, 0m, "Motivo");
        nc.Anular("USR01");

        Assert.Equal("AN", nc.Estado);
        Assert.Equal("USR01", nc.UsuarioAnulado);
        Assert.NotNull(nc.FechaAnulado);
    }

    [Fact]
    public void Emitir_SoloNetoSinImpuestos_TotalIgualNeto()
    {
        var nc = NotaCredito.Emitir("NC008", DateTime.Today, "DOC002", 200m, 0m, 0m, 0m, "Sin impuesto");

        Assert.Equal(200m, nc.Total);
    }
}
