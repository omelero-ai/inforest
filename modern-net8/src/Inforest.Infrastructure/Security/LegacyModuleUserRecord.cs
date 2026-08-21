namespace Inforest.Infrastructure.Security;

//internal sealed record LegacyModuleUserRecord(
//    string tCodigoUsuario,
//    string tGrupoUsuario,
//    string tDetallado,
//    string tResumido,
//    string tPassword,
//    bool lActivo,
//    string tBandaMagnetica);
internal sealed class LegacyModuleUserRecord
{
    public string tCodigoUsuario { get; init; } = string.Empty;
    public string? tGrupoUsuario { get; init; } = string.Empty;
    public string tDetallado { get; init; } = string.Empty;
    public string tResumido { get; init; } = string.Empty;
    public string tPassword { get; init; } = string.Empty;
    public bool lActivo { get; init; }
    public string? tBandaMagnetica { get; init; }
}
