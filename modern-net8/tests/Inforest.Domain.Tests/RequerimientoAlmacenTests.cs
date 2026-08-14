using Inforest.Domain.Entities.Almacen;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests;

/// <summary>
/// Tests de entidades de almacén — POS-FUNC-014.
/// BR-IMPORT-001, BR-IMPORT-002, BR-IMPORT-003, BR-IMPORT-004.
/// </summary>
public class RequerimientoAlmacenTests
{
    // ── RequerimientoAlmacen ─────────────────────────────────────────────────

    [Fact]
    public void Crear_DatosValidos_CreaRequerimientoCorrectamente()
    {
        var rq = RequerimientoAlmacen.Crear("REQ-001", DateTime.Today, "Cocina", "Juan Pérez", "02");

        Assert.Equal("REQ-001", rq.Rq);
        Assert.Equal("Cocina", rq.Area);
        Assert.Equal("Juan Pérez", rq.Responsable);
        Assert.Equal("02", rq.CodEstado);
        Assert.Null(rq.CodigoPedido);
        Assert.False(rq.YaImportado);
    }

    [Fact]
    public void Crear_CodigoVacio_LanzaExcepcion()
    {
        Assert.Throws<DomainException>(() =>
            RequerimientoAlmacen.Crear("", DateTime.Today, "Cocina", "Juan", "02"));
    }

    [Fact]
    public void PuedeImportarse_EstadoCorrecto_SinPedido_RetornaTrue_BR_IMPORT_001()
    {
        // BR-IMPORT-001: Estado '02' y sin pedido generado puede importarse
        var rq = RequerimientoAlmacen.Crear("REQ-001", DateTime.Today, "Cocina", "Juan", "02");
        Assert.True(rq.PuedeImportarse());
    }

    [Fact]
    public void PuedeImportarse_EstadoIncorrecto_RetornaFalse_BR_IMPORT_001()
    {
        // BR-IMPORT-001: Estado diferente a '02' no puede importarse
        var rq = RequerimientoAlmacen.Crear("REQ-001", DateTime.Today, "Cocina", "Juan", "01");
        Assert.False(rq.PuedeImportarse());
    }

    [Fact]
    public void YaImportado_ConPedido_RetornaTrue_BR_IMPORT_002()
    {
        // BR-IMPORT-002: Con pedido ya generado, está importado
        var rq = RequerimientoAlmacen.Crear("REQ-001", DateTime.Today, "Cocina", "Juan", "02", "PED-0001");
        Assert.True(rq.YaImportado);
        Assert.False(rq.PuedeImportarse());
    }

    [Fact]
    public void PuedeImportarse_ConPedidoAsociado_RetornaFalse_BR_IMPORT_002()
    {
        // BR-IMPORT-002: Un requerimiento que ya tiene pedido no puede volver a importarse
        var rq = RequerimientoAlmacen.Crear("REQ-001", DateTime.Today, "Cocina", "Juan", "02", "A001");
        Assert.False(rq.PuedeImportarse());
    }

    // ── DetalleRequerimientoAlmacen ──────────────────────────────────────────

    [Fact]
    public void Detalle_Crear_DatosValidos_CreaDetalleCorrectamente()
    {
        var detalle = DetalleRequerimientoAlmacen.Crear("ART-001", "Harina", 5.0, "PROD-01");

        Assert.Equal("ART-001", detalle.CodArt);
        Assert.Equal("Harina", detalle.Articulo);
        Assert.Equal(5.0, detalle.Cantidad);
        Assert.Equal("PROD-01", detalle.EnlaceInforRest);
        Assert.True(detalle.TieneEnlace);
    }

    [Fact]
    public void Detalle_SinEnlace_TieneEnlaceFalse_BR_IMPORT_003()
    {
        // BR-IMPORT-003: Sin enlace el artículo no puede importarse
        var detalle = DetalleRequerimientoAlmacen.Crear("ART-001", "Harina", 5.0);
        Assert.False(detalle.TieneEnlace);
        Assert.Null(detalle.EnlaceInforRest);
    }

    [Fact]
    public void Detalle_CantidadCero_LanzaExcepcion()
    {
        Assert.Throws<DomainException>(() =>
            DetalleRequerimientoAlmacen.Crear("ART-001", "Harina", 0));
    }

    [Fact]
    public void Detalle_CodArtVacio_LanzaExcepcion()
    {
        Assert.Throws<DomainException>(() =>
            DetalleRequerimientoAlmacen.Crear("", "Harina", 3.0));
    }
}
