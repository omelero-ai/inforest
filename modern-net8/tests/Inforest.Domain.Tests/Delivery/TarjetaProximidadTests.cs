using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Delivery;

/// <summary>
/// Tests de la entidad TarjetaProximidad.
/// BR-RFID-001: código/descripcion/cliente obligatorios y estado controlado.
/// BR-RFID-003: monto disponible no negativo.
/// </summary>
public sealed class TarjetaProximidadTests
{
    [Fact]
    public void Crear_CodigoVacio_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            TarjetaProximidad.Crear("", "Tarjeta VIP", 0m, "0001", "Free", "USR01"));

        Assert.Equal("RFID_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Crear_DescripcionVacia_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            TarjetaProximidad.Crear("RFID001", "", 0m, "0001", "Free", "USR01"));

        Assert.Equal("RFID_DESCRIPCION_REQUERIDA", ex.Code);
    }

    [Fact]
    public void Crear_ClienteVacio_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            TarjetaProximidad.Crear("RFID001", "Tarjeta VIP", 0m, "", "Free", "USR01"));

        Assert.Equal("RFID_CLIENTE_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Crear_EstadoInvalido_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            TarjetaProximidad.Crear("RFID001", "Tarjeta VIP", 0m, "0001", "Activa", "USR01"));

        Assert.Equal("RFID_ESTADO_INVALIDO", ex.Code);
    }

    [Fact]
    public void Crear_MontoNegativo_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() =>
            TarjetaProximidad.Crear("RFID001", "Tarjeta VIP", -1m, "0001", "Free", "USR01"));

        Assert.Equal("RFID_MONTO_INVALIDO", ex.Code);
    }

    [Fact]
    public void Crear_ValidaEstadoPermitido_NormalizaValor()
    {
        var tarjeta = TarjetaProximidad.Crear("RFID001", "Tarjeta VIP", 12.345m, "0001", "free", "USR01");

        Assert.Equal("RFID001", tarjeta.CodigoRfid);
        Assert.Equal("Free", tarjeta.Estado);
        Assert.Equal(12.35m, tarjeta.MontoDisponible);
    }

    [Fact]
    public void Actualizar_CamposValidos_ActualizaPropiedades()
    {
        var tarjeta = TarjetaProximidad.Crear("RFID001", "Tarjeta VIP", 5m, "0001", "Free", "USR01");

        tarjeta.Actualizar("Tarjeta Recargable", 15m, "0002", "Asignada", "USR02");

        Assert.Equal("Tarjeta Recargable", tarjeta.Descripcion);
        Assert.Equal("0002", tarjeta.CodigoCliente);
        Assert.Equal("Asignada", tarjeta.Estado);
        Assert.Equal(15m, tarjeta.MontoDisponible);
        Assert.Equal("USR02", tarjeta.UsuarioModificacion);
    }
}
