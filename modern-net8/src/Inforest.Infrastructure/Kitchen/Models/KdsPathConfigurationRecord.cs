namespace Inforest.Infrastructure.Kitchen.Models;

/// <summary>
/// Resultado de USP_KDS_ObtenerPath.
/// </summary>
internal sealed record KdsPathConfigurationRecord(
    string? tOrderInfo,
    string? tOrderStatus,
    string? tBump,
    string? tOrderInfox,
    string? tOrderStatusx,
    string? tBumpx,
    bool? lKDSInforest,
    bool? lKDSxArea,
    bool? lKDSTodos,
    bool? lKDS);
