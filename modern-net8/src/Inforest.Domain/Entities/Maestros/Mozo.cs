using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Maestro operativo de mozos/meseros.
/// Legacy: TMOZO / frmMozo.frm. BR-014.
/// </summary>
public class Mozo : Entity
{
    public string CodigoMozo { get; private set; } = string.Empty;
    public string Detallado { get; private set; } = string.Empty;
    public string Resumido { get; private set; } = string.Empty;
    public bool Activo { get; private set; }
    public string Usuario { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }

    private Mozo() { }

    public static Mozo Crear(
        string codigoMozo,
        string detallado,
        string? resumido,
        string usuario,
        bool activo = true,
        DateTime? fechaRegistro = null)
    {
        Validar(codigoMozo, detallado, usuario);

        return new Mozo
        {
            CodigoMozo = codigoMozo.Trim(),
            Detallado = detallado.Trim(),
            Resumido = NormalizarResumido(resumido, detallado),
            Usuario = usuario.Trim(),
            Activo = activo,
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };
    }

    public void Actualizar(string detallado, string? resumido, string usuario, bool activo)
    {
        Validar(CodigoMozo, detallado, usuario);
        Detallado = detallado.Trim();
        Resumido = NormalizarResumido(resumido, detallado);
        Usuario = usuario.Trim();
        Activo = activo;
    }

    public void Activar() => Activo = true;
    public void Inactivar() => Activo = false;

    private static void Validar(string codigoMozo, string detallado, string usuario)
    {
        if (string.IsNullOrWhiteSpace(codigoMozo))
            throw new DomainException("El código del mozo es obligatorio.", "MOZO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(detallado))
            throw new DomainException("La descripción del mozo es obligatoria.", "MOZO_DETALLADO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario del mozo es obligatorio.", "MOZO_USUARIO_REQUERIDO");
    }

    private static string NormalizarResumido(string? resumido, string detallado)
        => string.IsNullOrWhiteSpace(resumido)
            ? detallado.Trim()[..Math.Min(detallado.Trim().Length, 20)]
            : resumido.Trim();
}
