using System.Text.RegularExpressions;
using Dapper;
using Inforest.Application.Interfaces;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Security;

/// <summary>
/// Almacenamiento opcional de hashes BCrypt sin modificar TUSUARIO.tPassword (varchar(8)).
/// Requiere una tabla sidecar provisionada externamente.
/// </summary>
internal sealed class ModernPasswordHashStore : IModernPasswordHashStore
{
    private static readonly Regex SafeIdentifierRegex = new("^[A-Za-z0-9_]+$", RegexOptions.Compiled);

    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ILogger<ModernPasswordHashStore> _logger;
    private readonly string _tableName;

    public ModernPasswordHashStore(
        IDbConnectionFactory connectionFactory,
        IConfiguration configuration,
        ILogger<ModernPasswordHashStore> logger)
    {
        _connectionFactory = connectionFactory;
        _logger = logger;
        _tableName = configuration["Inforest:Security:ModernPasswordHashTable"] ?? "TUSUARIOPASSWORDHASH";
    }

    public async Task<string?> GetHashAsync(string codigoUsuario, CancellationToken cancellationToken = default)
    {
        if (!await TableExistsAsync(cancellationToken))
            return null;

        await using var connection = await _connectionFactory.CreateOpenConnectionAsync("INFSEGURIDAD", cancellationToken);
        var sql = $"SELECT tHashBCrypt FROM dbo.[{_tableName}] WHERE tCodigoUsuario = @CodigoUsuario";
        return await connection.QueryFirstOrDefaultAsync<string?>(new CommandDefinition(
            sql,
            new { CodigoUsuario = codigoUsuario },
            cancellationToken: cancellationToken));
    }

    public async Task<bool> UpsertHashAsync(string codigoUsuario, string hash, CancellationToken cancellationToken = default)
    {
        if (!await TableExistsAsync(cancellationToken))
            return false;

        await using var connection = await _connectionFactory.CreateOpenConnectionAsync("INFSEGURIDAD", cancellationToken);
        var sql = $"""
            MERGE dbo.[{_tableName}] AS target
            USING (SELECT @CodigoUsuario AS tCodigoUsuario, @Hash AS tHashBCrypt) AS source
                ON target.tCodigoUsuario = source.tCodigoUsuario
            WHEN MATCHED THEN
                UPDATE SET tHashBCrypt = source.tHashBCrypt, fActualizacion = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (tCodigoUsuario, tHashBCrypt, fActualizacion)
                VALUES (source.tCodigoUsuario, source.tHashBCrypt, GETDATE());
            """;

        await connection.ExecuteAsync(new CommandDefinition(
            sql,
            new { CodigoUsuario = codigoUsuario, Hash = hash },
            cancellationToken: cancellationToken));

        return true;
    }

    private async Task<bool> TableExistsAsync(CancellationToken cancellationToken)
    {
        if (!SafeIdentifierRegex.IsMatch(_tableName))
            return false;

        await using var connection = await _connectionFactory.CreateOpenConnectionAsync("INFSEGURIDAD", cancellationToken);
        const string sql = """
            SELECT COUNT(*)
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = @TableName
            """;

        var exists = await connection.ExecuteScalarAsync<int>(new CommandDefinition(
            sql,
            new { TableName = _tableName },
            cancellationToken: cancellationToken));

        if (exists == 0)
            _logger.LogDebug("Tabla sidecar de passwords {TableName} no configurada en INFSEGURIDAD", _tableName);

        return exists > 0;
    }
}
