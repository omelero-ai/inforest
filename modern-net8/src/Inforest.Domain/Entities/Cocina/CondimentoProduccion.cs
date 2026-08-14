using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Cocina;

/// <summary>
/// Condimento/modificador enviado a cocina.
/// Legacy: USP_KDS_ObtenerPropiedadesProducto + XML &lt;Condiment&gt; en modKDS.bas.
/// Regla BR-010.
/// </summary>
public sealed record CondimentoProduccion
{
    public string Codigo { get; }
    public string Nombre { get; }

    public CondimentoProduccion(string codigo, string nombre)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código del condimento es obligatorio.", "COCINA_CONDIMENTO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(nombre))
            throw new DomainException("El nombre del condimento es obligatorio.", "COCINA_CONDIMENTO_NOMBRE_REQUERIDO");

        Codigo = codigo.Trim();
        Nombre = nombre.Trim();
    }
}
