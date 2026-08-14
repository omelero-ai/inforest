using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Maestro operativo de mesas.
/// Legacy: TMESA / frmMesa.frm / frmMesas.frm. BR-003.
/// </summary>
public class Mesa : Entity
{
    public string CodigoMesa { get; private set; } = string.Empty;
    public string Detallado { get; private set; } = string.Empty;
    public string Resumido { get; private set; } = string.Empty;
    public string CodigoSalon { get; private set; } = string.Empty;
    public bool Fumador { get; private set; }
    public string PosicionX { get; private set; } = string.Empty;
    public string PosicionY { get; private set; } = string.Empty;
    public int NumeroPersonas { get; private set; }
    public bool Activo { get; private set; }
    public string Usuario { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }
    public EstadoMesa Estado { get; private set; }

    private Mesa() { }

    public static Mesa Crear(
        string codigoMesa,
        string detallado,
        string? resumido,
        string codigoSalon,
        bool fumador,
        string? posicionX,
        string? posicionY,
        int numeroPersonas,
        string usuario,
        EstadoMesa estado = EstadoMesa.Libre,
        bool activo = true,
        DateTime? fechaRegistro = null)
    {
        Validar(codigoMesa, detallado, codigoSalon, numeroPersonas, usuario);

        return new Mesa
        {
            CodigoMesa = codigoMesa.Trim(),
            Detallado = detallado.Trim(),
            Resumido = NormalizarResumido(resumido, detallado),
            CodigoSalon = codigoSalon.Trim(),
            Fumador = fumador,
            PosicionX = string.IsNullOrWhiteSpace(posicionX) ? "0" : posicionX.Trim(),
            PosicionY = string.IsNullOrWhiteSpace(posicionY) ? "0" : posicionY.Trim(),
            NumeroPersonas = numeroPersonas,
            Usuario = usuario.Trim(),
            Estado = estado,
            Activo = activo,
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };
    }

    public void Actualizar(
        string detallado,
        string? resumido,
        string codigoSalon,
        bool fumador,
        string? posicionX,
        string? posicionY,
        int numeroPersonas,
        string usuario,
        bool activo)
    {
        Validar(CodigoMesa, detallado, codigoSalon, numeroPersonas, usuario);
        Detallado = detallado.Trim();
        Resumido = NormalizarResumido(resumido, detallado);
        CodigoSalon = codigoSalon.Trim();
        Fumador = fumador;
        PosicionX = string.IsNullOrWhiteSpace(posicionX) ? "0" : posicionX.Trim();
        PosicionY = string.IsNullOrWhiteSpace(posicionY) ? "0" : posicionY.Trim();
        NumeroPersonas = numeroPersonas;
        Usuario = usuario.Trim();
        Activo = activo;
    }

    public void Ocupar()
    {
        if (Estado == EstadoMesa.Ocupada)
            throw new DomainException("La mesa ya se encuentra ocupada.", "MESA_YA_OCUPADA");

        Estado = EstadoMesa.Ocupada;
    }

    public void Reservar()
    {
        if (Estado == EstadoMesa.Ocupada)
            throw new DomainException("No se puede reservar una mesa ocupada.", "MESA_OCUPADA_NO_RESERVABLE");

        Estado = EstadoMesa.Reservada;
    }

    public void MarcarEnCuenta()
    {
        if (Estado == EstadoMesa.Libre)
            throw new DomainException("Una mesa libre no puede pasar directamente a en cuenta.", "MESA_LIBRE_NO_EN_CUENTA");

        Estado = EstadoMesa.EnCuenta;
    }

    public void Liberar() => Estado = EstadoMesa.Libre;
    public void Activar() => Activo = true;
    public void Inactivar() => Activo = false;

    private static void Validar(string codigoMesa, string detallado, string codigoSalon, int numeroPersonas, string usuario)
    {
        if (string.IsNullOrWhiteSpace(codigoMesa))
            throw new DomainException("El código de mesa es obligatorio.", "MESA_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(detallado))
            throw new DomainException("La descripción de la mesa es obligatoria.", "MESA_DETALLADO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoSalon))
            throw new DomainException("El salón de la mesa es obligatorio.", "MESA_SALON_REQUERIDO");

        if (numeroPersonas <= 0)
            throw new DomainException("La capacidad de la mesa debe ser mayor a cero.", "MESA_CAPACIDAD_INVALIDA");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario de la mesa es obligatorio.", "MESA_USUARIO_REQUERIDO");
    }

    private static string NormalizarResumido(string? resumido, string detallado)
        => string.IsNullOrWhiteSpace(resumido)
            ? detallado.Trim()[..Math.Min(detallado.Trim().Length, 20)]
            : resumido.Trim();
}
