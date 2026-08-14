using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Insumo o plato de stock crítico.
/// Legacy: TINSUMO / frmInsumo.frm / frmInsumoDetalle.frm.
/// BR-INSUMO-001, BR-INSUMO-002, BR-INSUMO-003, BR-INSUMO-004.
/// </summary>
public class Insumo : Entity
{
    public string Codigo { get; private set; } = string.Empty;
    public string Usuario { get; private set; } = string.Empty;
    public string Descripcion { get; private set; } = string.Empty;
    public double Stock { get; private set; }
    public DateTime FechaRegistro { get; private set; }
    public bool Activo { get; private set; }

    /// <summary>
    /// true = Insumo de almacén; false = Plato/producto de stock crítico.
    /// Legacy: lINSUMO en TINSUMO.
    /// </summary>
    public bool EsInsumo { get; private set; }

    private Insumo() { }

    /// <summary>
    /// Crea un nuevo insumo.
    /// BR-INSUMO-001: La descripción no puede estar vacía.
    /// BR-INSUMO-002: La descripción se almacena en mayúsculas.
    /// </summary>
    public static Insumo Crear(
        string codigo,
        string descripcion,
        string usuario,
        double stock = 0,
        bool esInsumo = true,
        bool activo = true,
        DateTime? fechaRegistro = null)
    {
        ValidarCreacion(codigo, descripcion, usuario);

        return new Insumo
        {
            Codigo = codigo.Trim(),
            Descripcion = descripcion.Trim().ToUpperInvariant(),
            Usuario = usuario.Trim(),
            Stock = stock,
            EsInsumo = esInsumo,
            Activo = activo,
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };
    }

    /// <summary>
    /// Actualiza los datos del insumo.
    /// BR-INSUMO-003: Solo se puede modificar un insumo activo o explícitamente para darlo de baja.
    /// </summary>
    public void Actualizar(
        string descripcion,
        string usuario,
        double stock,
        bool esInsumo,
        bool activo)
    {
        if (string.IsNullOrWhiteSpace(descripcion))
            throw new DomainException("La descripción del insumo no puede estar vacía.", "INSUMO_DESCRIPCION_REQUERIDA");
        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario es requerido.", "INSUMO_USUARIO_REQUERIDO");

        Descripcion = descripcion.Trim().ToUpperInvariant();
        Usuario = usuario.Trim();
        Stock = stock;
        EsInsumo = esInsumo;
        Activo = activo;
    }

    private static void ValidarCreacion(string codigo, string descripcion, string usuario)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código del insumo no puede estar vacío.", "INSUMO_CODIGO_REQUERIDO");
        if (string.IsNullOrWhiteSpace(descripcion))
            throw new DomainException("La descripción del insumo no puede estar vacía.", "INSUMO_DESCRIPCION_REQUERIDA");
        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario es requerido.", "INSUMO_USUARIO_REQUERIDO");
    }
}
