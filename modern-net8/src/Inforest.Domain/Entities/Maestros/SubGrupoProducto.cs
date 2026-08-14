using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Maestro de subgrupos de producto.
/// Legacy: TSUBGRUPOPRODUCTO / frmGrupo.frm. BR-014.
/// </summary>
public class SubGrupoProducto : Entity
{
    public string CodigoSubGrupo { get; private set; } = string.Empty;
    public string CodigoGrupo { get; private set; } = string.Empty;
    public string Detallado { get; private set; } = string.Empty;
    public string Resumido { get; private set; } = string.Empty;
    public bool Activo { get; private set; }
    public string Usuario { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }

    private SubGrupoProducto() { }

    public static SubGrupoProducto Crear(
        string codigoSubGrupo,
        string codigoGrupo,
        string detallado,
        string? resumido,
        string usuario,
        bool activo = true,
        DateTime? fechaRegistro = null)
    {
        Validar(codigoSubGrupo, codigoGrupo, detallado, usuario);

        return new SubGrupoProducto
        {
            CodigoSubGrupo = codigoSubGrupo.Trim(),
            CodigoGrupo = codigoGrupo.Trim(),
            Detallado = detallado.Trim(),
            Resumido = NormalizarResumido(resumido, detallado),
            Usuario = usuario.Trim(),
            Activo = activo,
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };
    }

    public void Actualizar(string codigoGrupo, string detallado, string? resumido, string usuario, bool activo)
    {
        Validar(CodigoSubGrupo, codigoGrupo, detallado, usuario);
        CodigoGrupo = codigoGrupo.Trim();
        Detallado = detallado.Trim();
        Resumido = NormalizarResumido(resumido, detallado);
        Usuario = usuario.Trim();
        Activo = activo;
    }

    public void Activar() => Activo = true;
    public void Inactivar() => Activo = false;

    private static void Validar(string codigoSubGrupo, string codigoGrupo, string detallado, string usuario)
    {
        if (string.IsNullOrWhiteSpace(codigoSubGrupo))
            throw new DomainException("El código del subgrupo es obligatorio.", "SUBGRUPO_PRODUCTO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoGrupo))
            throw new DomainException("El grupo padre del subgrupo es obligatorio.", "SUBGRUPO_PRODUCTO_GRUPO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(detallado))
            throw new DomainException("La descripción detallada del subgrupo es obligatoria.", "SUBGRUPO_PRODUCTO_DETALLADO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario de registro del subgrupo es obligatorio.", "SUBGRUPO_PRODUCTO_USUARIO_REQUERIDO");
    }

    private static string NormalizarResumido(string? resumido, string detallado)
        => string.IsNullOrWhiteSpace(resumido)
            ? detallado.Trim()[..Math.Min(detallado.Trim().Length, 20)]
            : resumido.Trim();
}
