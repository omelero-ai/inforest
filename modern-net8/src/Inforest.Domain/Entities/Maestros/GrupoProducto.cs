using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Maestro de grupos de producto.
/// Legacy: TGRUPOPRODUCTO / frmGrupo.frm. BR-014.
/// </summary>
public class GrupoProducto : Entity
{
    public string CodigoGrupo { get; private set; } = string.Empty;
    public string Detallado { get; private set; } = string.Empty;
    public string Resumido { get; private set; } = string.Empty;
    public bool Activo { get; private set; }
    public string Usuario { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }

    private GrupoProducto() { }

    public static GrupoProducto Crear(
        string codigoGrupo,
        string detallado,
        string? resumido,
        string usuario,
        bool activo = true,
        DateTime? fechaRegistro = null)
    {
        Validar(codigoGrupo, detallado, usuario);

        return new GrupoProducto
        {
            CodigoGrupo = codigoGrupo.Trim(),
            Detallado = detallado.Trim(),
            Resumido = NormalizarResumido(resumido, detallado),
            Usuario = usuario.Trim(),
            Activo = activo,
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };
    }

    public void Actualizar(string detallado, string? resumido, string usuario, bool activo)
    {
        Validar(CodigoGrupo, detallado, usuario);
        Detallado = detallado.Trim();
        Resumido = NormalizarResumido(resumido, detallado);
        Usuario = usuario.Trim();
        Activo = activo;
    }

    public void Activar() => Activo = true;
    public void Inactivar() => Activo = false;

    private static void Validar(string codigoGrupo, string detallado, string usuario)
    {
        if (string.IsNullOrWhiteSpace(codigoGrupo))
            throw new DomainException("El código del grupo es obligatorio.", "GRUPO_PRODUCTO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(detallado))
            throw new DomainException("La descripción detallada del grupo es obligatoria.", "GRUPO_PRODUCTO_DETALLADO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario de registro del grupo es obligatorio.", "GRUPO_PRODUCTO_USUARIO_REQUERIDO");
    }

    private static string NormalizarResumido(string? resumido, string detallado)
        => string.IsNullOrWhiteSpace(resumido)
            ? detallado.Trim()[..Math.Min(detallado.Trim().Length, 20)]
            : resumido.Trim();
}
