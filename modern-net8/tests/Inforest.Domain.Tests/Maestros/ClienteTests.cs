using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Maestros;

/// <summary>
/// Tests del maestro de clientes.
/// Legacy: TCLIENTE / frmCliente.frm.
/// </summary>
public class ClienteTests
{
    [Fact]
    public void Crear_ClienteValido_CreaEntidad()
    {
        var cliente = Cliente.Crear("CLI001", "Cliente Demo", "20123456789", "Av. Principal 123", "tester", tipoIdentidad: "RUC");

        Assert.Equal("CLI001", cliente.CodigoCliente);
        Assert.Equal("20123456789", cliente.Identidad);
    }

    [Fact]
    public void Crear_RucInvalido_LanzaDomainException()
    {
        var ex = Assert.Throws<DomainException>(() =>
            Cliente.Crear("CLI001", "Cliente Demo", "123", "Av. Principal 123", "tester", tipoIdentidad: "RUC"));

        Assert.Equal("CLIENTE_RUC_INVALIDO", ex.Code);
    }

    [Fact]
    public void Crear_DniInvalido_LanzaDomainException()
    {
        var ex = Assert.Throws<DomainException>(() =>
            Cliente.Crear("CLI002", "Cliente Demo", "1234567", "Av. Principal 123", "tester", tipoIdentidad: "DNI"));

        Assert.Equal("CLIENTE_DNI_INVALIDO", ex.Code);
    }
}
