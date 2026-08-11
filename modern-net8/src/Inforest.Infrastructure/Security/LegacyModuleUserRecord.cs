namespace Inforest.Infrastructure.Security;

internal sealed record LegacyModuleUserRecord(
    string tCodigoUsuario,
    string tGrupoUsuario,
    string tDetallado,
    string tResumido,
    string tPassword,
    bool lActivo,
    string tBandaMagnetica);
