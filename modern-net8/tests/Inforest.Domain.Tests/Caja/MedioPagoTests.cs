using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Caja;

public class MedioPagoTests
{
    [Fact]
    public void Crear_MedioPagoValido_RetornaEntidadActiva()
    {
        var medioPago = MedioPago.Crear("01", "Efectivo", TipoMedioPago.Efectivo);

        Assert.Equal("01", medioPago.Codigo);
        Assert.Equal("Efectivo", medioPago.Descripcion);
        Assert.True(medioPago.Activo);
    }

    [Fact]
    public void CrearPagoDocumento_ConMontoNegativo_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            PagoDocumento.Crear("DOC001", "01", -10m, 0m, null));

        Assert.Equal("PAGO_MONTO_INVALIDO", ex.Code);
    }
}
