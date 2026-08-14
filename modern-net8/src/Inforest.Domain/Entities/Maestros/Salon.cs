using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Salón operativo donde se agrupan mesas.
/// Legacy: TSALON / TTABLA(SALON) / frmMesa.frm. BR-003.
/// </summary>
public class Salon : Entity
{
    public string CodigoSalon { get; private set; } = string.Empty;
    public string Detallado { get; private set; } = string.Empty;
    public string Resumido { get; private set; } = string.Empty;
    public bool Activo { get; private set; }
    public string Usuario { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }

    private Salon() { }

    public static Salon Crear(
        string codigoSalon,
        string detallado,
        string? resumido,
        string usuario,
        bool activo = true,
        DateTime? fechaRegistro = null)
    {
        Validar(codigoSalon, detallado, usuario);

        return new Salon
        {
            CodigoSalon = codigoSalon.Trim(),
            Detallado = detallado.Trim(),
            Resumido = NormalizarResumido(resumido, detallado),
            Usuario = usuario.Trim(),
            Activo = activo,
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };
    }

    public void Actualizar(string detallado, string? resumido, string usuario, bool activo)
    {
        Validar(CodigoSalon, detallado, usuario);
        Detallado = detallado.Trim();
        Resumido = NormalizarResumido(resumido, detallado);
        Usuario = usuario.Trim();
        Activo = activo;
    }

    public void Activar() => Activo = true;
    public void Inactivar() => Activo = false;

    private static void Validar(string codigoSalon, string detallado, string usuario)
    {
        if (string.IsNullOrWhiteSpace(codigoSalon))
            throw new DomainException("El código del salón es obligatorio.", "SALON_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(detallado))
            throw new DomainException("La descripción del salón es obligatoria.", "SALON_DETALLADO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario del salón es obligatorio.", "SALON_USUARIO_REQUERIDO");
    }

    private static string NormalizarResumido(string? resumido, string detallado)
        => string.IsNullOrWhiteSpace(resumido)
            ? detallado.Trim()[..Math.Min(detallado.Trim().Length, 20)]
            : resumido.Trim();
}
