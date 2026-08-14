using System.Text.RegularExpressions;
using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Maestro de clientes del restaurante.
/// Legacy: TCLIENTE / frmCliente.frm. BR-014.
/// </summary>
public class Cliente : Entity
{
    public string CodigoCliente { get; private set; } = string.Empty;
    public string Empresa { get; private set; } = string.Empty;
    public string Identidad { get; private set; } = string.Empty;
    public string Direccion { get; private set; } = string.Empty;
    public bool Activo { get; private set; }
    public string Usuario { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }
    public string? Correo { get; private set; }
    public string? TipoIdentidad { get; private set; }
    public string? Enlace { get; private set; }
    public string? TipoCliente { get; private set; }
    public string? Ubigeo { get; private set; }
    public string? Urbanizacion { get; private set; }
    public string? TipoResponsable { get; private set; }

    private Cliente() { }

    public static Cliente Crear(
        string codigoCliente,
        string empresa,
        string identidad,
        string direccion,
        string usuario,
        string? correo = null,
        string? tipoIdentidad = null,
        string? enlace = null,
        string? tipoCliente = null,
        string? ubigeo = null,
        string? urbanizacion = null,
        string? tipoResponsable = null,
        bool activo = true,
        DateTime? fechaRegistro = null)
    {
        Validar(codigoCliente, empresa, identidad, direccion, usuario, tipoIdentidad, correo);

        return new Cliente
        {
            CodigoCliente = codigoCliente.Trim(),
            Empresa = empresa.Trim(),
            Identidad = identidad.Trim(),
            Direccion = direccion.Trim(),
            Usuario = usuario.Trim(),
            Correo = string.IsNullOrWhiteSpace(correo) ? null : correo.Trim(),
            TipoIdentidad = string.IsNullOrWhiteSpace(tipoIdentidad) ? null : tipoIdentidad.Trim(),
            Enlace = string.IsNullOrWhiteSpace(enlace) ? null : enlace.Trim(),
            TipoCliente = string.IsNullOrWhiteSpace(tipoCliente) ? null : tipoCliente.Trim(),
            Ubigeo = string.IsNullOrWhiteSpace(ubigeo) ? null : ubigeo.Trim(),
            Urbanizacion = string.IsNullOrWhiteSpace(urbanizacion) ? null : urbanizacion.Trim(),
            TipoResponsable = string.IsNullOrWhiteSpace(tipoResponsable) ? null : tipoResponsable.Trim(),
            Activo = activo,
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };
    }

    public void Actualizar(
        string empresa,
        string identidad,
        string direccion,
        string usuario,
        string? correo,
        string? tipoIdentidad,
        string? enlace,
        string? tipoCliente,
        string? ubigeo,
        string? urbanizacion,
        string? tipoResponsable,
        bool activo)
    {
        Validar(CodigoCliente, empresa, identidad, direccion, usuario, tipoIdentidad, correo);

        Empresa = empresa.Trim();
        Identidad = identidad.Trim();
        Direccion = direccion.Trim();
        Usuario = usuario.Trim();
        Correo = string.IsNullOrWhiteSpace(correo) ? null : correo.Trim();
        TipoIdentidad = string.IsNullOrWhiteSpace(tipoIdentidad) ? null : tipoIdentidad.Trim();
        Enlace = string.IsNullOrWhiteSpace(enlace) ? null : enlace.Trim();
        TipoCliente = string.IsNullOrWhiteSpace(tipoCliente) ? null : tipoCliente.Trim();
        Ubigeo = string.IsNullOrWhiteSpace(ubigeo) ? null : ubigeo.Trim();
        Urbanizacion = string.IsNullOrWhiteSpace(urbanizacion) ? null : urbanizacion.Trim();
        TipoResponsable = string.IsNullOrWhiteSpace(tipoResponsable) ? null : tipoResponsable.Trim();
        Activo = activo;
    }

    public void Activar() => Activo = true;
    public void Inactivar() => Activo = false;

    private static void Validar(
        string codigoCliente,
        string empresa,
        string identidad,
        string direccion,
        string usuario,
        string? tipoIdentidad,
        string? correo)
    {
        if (string.IsNullOrWhiteSpace(codigoCliente))
            throw new DomainException("El código del cliente es obligatorio.", "CLIENTE_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(empresa))
            throw new DomainException("La razón social o nombre del cliente es obligatoria.", "CLIENTE_EMPRESA_REQUERIDA");

        if (string.IsNullOrWhiteSpace(identidad))
            throw new DomainException("La identidad del cliente es obligatoria.", "CLIENTE_IDENTIDAD_REQUERIDA");

        if (string.IsNullOrWhiteSpace(direccion))
            throw new DomainException("La dirección del cliente es obligatoria.", "CLIENTE_DIRECCION_REQUERIDA");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario del cliente es obligatorio.", "CLIENTE_USUARIO_REQUERIDO");

        ValidarIdentidad(tipoIdentidad, identidad);

        if (!string.IsNullOrWhiteSpace(correo) && !Regex.IsMatch(correo, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            throw new DomainException("El correo del cliente no tiene un formato válido.", "CLIENTE_CORREO_INVALIDO");
    }

    private static void ValidarIdentidad(string? tipoIdentidad, string identidad)
    {
        var valor = identidad.Trim();
        var tipo = (tipoIdentidad ?? string.Empty).Trim().ToUpperInvariant();

        if (tipo is "06" or "RUC")
        {
            if (!Regex.IsMatch(valor, @"^\d{11}$"))
                throw new DomainException("El RUC del cliente debe tener 11 dígitos.", "CLIENTE_RUC_INVALIDO");

            return;
        }

        if (tipo is "01" or "DNI")
        {
            if (!Regex.IsMatch(valor, @"^\d{8}$"))
                throw new DomainException("El DNI del cliente debe tener 8 dígitos.", "CLIENTE_DNI_INVALIDO");

            return;
        }

        if (valor.Length < 4)
            throw new DomainException("La identidad del cliente debe tener al menos 4 caracteres.", "CLIENTE_IDENTIDAD_INVALIDA");
    }
}
