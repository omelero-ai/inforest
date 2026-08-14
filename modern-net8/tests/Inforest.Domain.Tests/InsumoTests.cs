using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests;

/// <summary>
/// Tests de entidad Insumo — POS-FUNC-013.
/// BR-INSUMO-001, BR-INSUMO-002, BR-INSUMO-003.
/// </summary>
public class InsumoTests
{
    [Fact]
    public void Crear_DatosValidos_CreaInsumoCorrectamente()
    {
        var insumo = Insumo.Crear("00000001", "sal", "user", 5.0, true);

        Assert.Equal("00000001", insumo.Codigo);
        Assert.Equal("user", insumo.Usuario);
        Assert.Equal(5.0, insumo.Stock);
        Assert.True(insumo.EsInsumo);
        Assert.True(insumo.Activo);
    }

    [Fact]
    public void Crear_DescripcionEnMayusculas_BR_INSUMO_002()
    {
        // BR-INSUMO-002: la descripción se almacena en mayúsculas
        var insumo = Insumo.Crear("00000001", "aceite de oliva", "user");
        Assert.Equal("ACEITE DE OLIVA", insumo.Descripcion);
    }

    [Fact]
    public void Crear_CodigoVacio_LanzaExcepcion_BR_INSUMO_001()
    {
        // BR-INSUMO-001: código requerido
        Assert.Throws<DomainException>(() => Insumo.Crear("", "SAL", "user"));
    }

    [Fact]
    public void Crear_DescripcionVacia_LanzaExcepcion_BR_INSUMO_001()
    {
        // BR-INSUMO-001: descripción requerida
        Assert.Throws<DomainException>(() => Insumo.Crear("00000001", "", "user"));
    }

    [Fact]
    public void Crear_UsuarioVacio_LanzaExcepcion()
    {
        Assert.Throws<DomainException>(() => Insumo.Crear("00000001", "SAL", ""));
    }

    [Fact]
    public void Actualizar_DatosValidos_ActualizaCorrectamente_BR_INSUMO_003()
    {
        // BR-INSUMO-003: modificación aplica nuevos valores
        var insumo = Insumo.Crear("00000001", "sal", "user", 5);
        insumo.Actualizar("sal gruesa", "user2", 15, true, false);

        Assert.Equal("SAL GRUESA", insumo.Descripcion);
        Assert.Equal("user2", insumo.Usuario);
        Assert.Equal(15, insumo.Stock);
        Assert.False(insumo.Activo);
    }

    [Fact]
    public void Actualizar_DescripcionVacia_LanzaExcepcion()
    {
        var insumo = Insumo.Crear("00000001", "sal", "user");
        Assert.Throws<DomainException>(() => insumo.Actualizar("", "user", 0, true, true));
    }

    [Fact]
    public void Crear_EsPlato_EsInsumoFalse()
    {
        var plato = Insumo.Crear("00000001", "LOMO SALTADO", "user", esInsumo: false);
        Assert.False(plato.EsInsumo);
    }

    [Fact]
    public void Crear_StockPredeterminado_EsCero()
    {
        var insumo = Insumo.Crear("00000001", "SAL", "user");
        Assert.Equal(0, insumo.Stock);
    }
}
