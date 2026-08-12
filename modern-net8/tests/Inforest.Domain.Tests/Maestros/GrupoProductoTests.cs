using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Maestros;

/// <summary>
/// Tests de grupo de producto.
/// Legacy: TGRUPOPRODUCTO / frmGrupo.frm.
/// </summary>
public class GrupoProductoTests
{
    [Fact]
    public void Crear_Valido_CreaEntidadActiva()
    {
        var grupo = GrupoProducto.Crear("BEB", "Bebidas", "BEBIDAS", "tester");

        Assert.Equal("BEB", grupo.CodigoGrupo);
        Assert.True(grupo.Activo);
    }

    [Fact]
    public void Crear_CodigoVacio_LanzaDomainException()
    {
        var ex = Assert.Throws<DomainException>(() => GrupoProducto.Crear("", "Bebidas", null, "tester"));
        Assert.Equal("GRUPO_PRODUCTO_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void InactivarYActivar_CambiaEstado()
    {
        var grupo = GrupoProducto.Crear("BEB", "Bebidas", null, "tester");

        grupo.Inactivar();
        Assert.False(grupo.Activo);

        grupo.Activar();
        Assert.True(grupo.Activo);
    }
}
