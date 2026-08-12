namespace Inforest.Infrastructure.Kitchen.Models;

/// <summary>
/// Resultado de USP_KDS_ObtenerPropiedadesProducto.
/// </summary>
internal sealed record KdsPropiedadProductoRecord(
    string tCodigoPropiedad,
    string? Operador,
    string? Propiedad,
    string? tEnlace,
    string? tProducto);
