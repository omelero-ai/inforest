namespace Inforest.Infrastructure.Security;

internal interface IModernPasswordHashStore
{
    Task<string?> GetHashAsync(string codigoUsuario, CancellationToken cancellationToken = default);

    Task<bool> UpsertHashAsync(string codigoUsuario, string hash, CancellationToken cancellationToken = default);
}
