namespace Inforest.Infrastructure.Kitchen.Models;

/// <summary>
/// Resultado de USP_KDS_ObtenerAreaImpresionKDS.
/// </summary>
internal sealed record KdsAreaImpresionRecord(
    string tCodigoProducto,
    string tArea,
    int KDSx);
