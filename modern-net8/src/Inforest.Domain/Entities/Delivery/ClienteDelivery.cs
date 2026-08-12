using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Delivery;

/// <summary>
/// Cliente de delivery registrado en el sistema.
/// <para>
/// Legacy: tabla <c>TDELIVERY</c> — cabecera del cliente con datos de contacto, zona y distrito.
/// </para>
/// Reglas: BR-DEL-001, BR-DEL-005
/// </summary>
public class ClienteDelivery : Entity
{
    public string CodigoDelivery { get; private set; } = string.Empty;
    public string? TipoCliente { get; private set; }
    public string? Apellido { get; private set; }
    public string? Nombre { get; private set; }
    public string? Direccion { get; private set; }
    public string? Telefono { get; private set; }
    public string? Referencia { get; private set; }
    public string? CodigoZona { get; private set; }
    public string? CodigoDistrito { get; private set; }
    public string? CodigoCliente { get; private set; }
    public string? CodigoTarjeta { get; private set; }
    public string? NumeroTarjeta { get; private set; }
    public string? TipoIdentidad { get; private set; }
    public string? NumeroIdentidad { get; private set; }
    public string? Observacion { get; private set; }
    public string? Email { get; private set; }
    public DateTime? FechaNacimiento { get; private set; }
    public bool Activo { get; private set; }

    // Dapper constructor
    private ClienteDelivery() { }

    public static ClienteDelivery Crear(
        string codigoDelivery,
        string? tipoCliente,
        string? apellido,
        string? nombre,
        string? telefono,
        string? direccion,
        string? codigoZona,
        string? codigoDistrito)
    {
        if (string.IsNullOrWhiteSpace(codigoDelivery))
            throw new DomainException("El código de delivery es obligatorio.", "DELIVERY_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(telefono) && string.IsNullOrWhiteSpace(apellido))
            throw new DomainException("El cliente debe tener al menos teléfono o apellido.", "DELIVERY_CONTACTO_REQUERIDO");

        return new ClienteDelivery
        {
            CodigoDelivery = codigoDelivery,
            TipoCliente = tipoCliente,
            Apellido = apellido,
            Nombre = nombre,
            Telefono = telefono,
            Direccion = direccion,
            CodigoZona = codigoZona,
            CodigoDistrito = codigoDistrito,
            Activo = true
        };
    }

    public void Actualizar(
        string? apellido,
        string? nombre,
        string? telefono,
        string? direccion,
        string? referencia,
        string? codigoZona,
        string? codigoDistrito,
        string? observacion,
        string? email)
    {
        Apellido = apellido;
        Nombre = nombre;
        Telefono = telefono;
        Direccion = direccion;
        Referencia = referencia;
        CodigoZona = codigoZona;
        CodigoDistrito = codigoDistrito;
        Observacion = observacion;
        Email = email;
    }

    public void AsignarIdentidad(string tipoIdentidad, string numeroIdentidad)
    {
        TipoIdentidad = tipoIdentidad;
        NumeroIdentidad = numeroIdentidad;
    }

    public void AsignarTarjeta(string codigoTarjeta, string numeroTarjeta)
    {
        CodigoTarjeta = codigoTarjeta;
        NumeroTarjeta = numeroTarjeta;
    }
}
