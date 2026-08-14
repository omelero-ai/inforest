using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Maestros;

/// <summary>
/// Tests de mesas operativas.
/// Legacy: TMESA / frmMesa.frm.
/// </summary>
public class MesaTests
{
    [Fact]
    public void Ocupar_Y_MarcarEnCuenta_CambiaEstado()
    {
        var mesa = Mesa.Crear("M01", "Mesa 1", "M1", "SAL01", false, null, null, 4, "tester");

        mesa.Ocupar();
        mesa.MarcarEnCuenta();

        Assert.Equal(EstadoMesa.EnCuenta, mesa.Estado);
    }

    [Fact]
    public void MesaOcupada_NoPuedeOcuparseDeNuevo()
    {
        var mesa = Mesa.Crear("M01", "Mesa 1", "M1", "SAL01", false, null, null, 4, "tester");
        mesa.Ocupar();

        var ex = Assert.Throws<DomainException>(() => mesa.Ocupar());
        Assert.Equal("MESA_YA_OCUPADA", ex.Code);
    }
}
