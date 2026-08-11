using System.Collections.Concurrent;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Seguridad;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Security;

/// <summary>
/// RBAC en memoria sobre tablas Legacy.
/// Legacy: TUSUARIO/TGRUPOACCESO/TACCESO.
/// </summary>
internal sealed class RbacService : IRbacService
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ILogger<RbacService> _logger;
    private readonly ConcurrentDictionary<string, IReadOnlyCollection<PermisoAcceso>> _cache = new();

    public RbacService(IDbConnectionFactory connectionFactory, ILogger<RbacService> logger)
    {
        _connectionFactory = connectionFactory;
        _logger = logger;
    }

    public async Task<IReadOnlyCollection<PermisoAcceso>> ObtenerPermisosAsync(string codigoUsuario, string modulo, CancellationToken cancellationToken = default)
    {
        var normalizedUser = Normalize(codigoUsuario);
        var normalizedModule = Normalize(modulo);
        var cacheKey = $"{normalizedUser}:{normalizedModule}";

        if (_cache.TryGetValue(cacheKey, out var cached))
            return cached;

        await using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            SELECT
                a.tCodigoAcceso AS CodigoAcceso,
                ISNULL(a.tModulo, '') AS Modulo,
                ISNULL(a.tFormulario, '') AS Formulario,
                ISNULL(a.tNombreObjeto, '') AS NombreObjeto,
                ISNULL(a.tDescripcion, '') AS Descripcion
            FROM dbo.TUSUARIO u
            INNER JOIN dbo.TGRUPOACCESO ga
                ON ga.tGrupoUsuario = u.tGrupoUsuario
            INNER JOIN dbo.TACCESO a
                ON a.tCodigoAcceso = ga.tCodigoAcceso
            WHERE UPPER(u.tResumido) = @CodigoUsuario
              AND ISNULL(u.lActivo, 0) = 1
              AND UPPER(ISNULL(a.tModulo, '')) = @Modulo
            ORDER BY ISNULL(a.nOrden, 0), a.tCodigoAcceso
            """;

        var records = await connection.QueryAsync<PermissionRecord>(new CommandDefinition(
            sql,
            new { CodigoUsuario = normalizedUser, Modulo = normalizedModule },
            cancellationToken: cancellationToken));

        var permissions = records
            .Select(record => new PermisoAcceso(
                record.CodigoAcceso,
                record.Modulo,
                record.Formulario,
                record.NombreObjeto,
                record.Descripcion))
            .ToArray();

        _cache[cacheKey] = permissions;
        _logger.LogDebug("Permisos RBAC cargados para {Usuario}/{Modulo}: {Cantidad}", normalizedUser, normalizedModule, permissions.Length);
        return permissions;
    }

    public async Task<bool> TieneAccesoAsync(string codigoUsuario, string modulo, string codigoAcceso, CancellationToken cancellationToken = default)
    {
        var permissions = await ObtenerPermisosAsync(codigoUsuario, modulo, cancellationToken);
        return permissions.Any(permission => string.Equals(permission.CodigoAcceso, codigoAcceso, StringComparison.OrdinalIgnoreCase));
    }

    private static string Normalize(string value) => value.Trim().ToUpperInvariant();

    private sealed record PermissionRecord(
        string CodigoAcceso,
        string Modulo,
        string Formulario,
        string NombreObjeto,
        string Descripcion);
}
