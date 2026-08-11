using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Infrastructure.Exceptions;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Data;

/// <summary>
/// Ejecutor de stored procedures Legacy usando Dapper.
/// <para>
/// Equivalencia con <c>clsComando.cls</c>:
/// </para>
/// <list type="table">
///   <listheader><term>Legacy</term><description>.NET 8</description></listheader>
///   <item><term>CreateCmdSp / CommandTimeout = 600</term><description>CommandTimeout = 600 en cada llamada Dapper</description></item>
///   <item><term>CreateParameter / Cmd.Parameters.Append</term><description>Objeto anónimo o <see cref="DynamicParameters"/> pasado en <c>parameters</c></description></item>
///   <item><term>ExecSP()</term><description><see cref="ExecuteAsync"/></description></item>
///   <item><term>GetSP() + iteración</term><description><see cref="QueryAsync{T}"/></description></item>
///   <item><term>GetSP() primer registro</term><description><see cref="QueryFirstOrDefaultAsync{T}"/></description></item>
///   <item><term>GetParameterValue</term><description><see cref="DynamicParameters"/> con <c>ParameterDirection.Output</c></description></item>
///   <item><term>ParamOK</term><description>Excepción tipificada <see cref="InfrastructureException"/> en caso de error</description></item>
/// </list>
/// Regla BR-SQL-CMD-001: timeout de 600 segundos heredado del Legacy.
/// </summary>
internal sealed class SpExecutor : ISpExecutor
{
    /// <summary>
    /// Timeout en segundos heredado de <c>clsComando.cls → Cmd.CommandTimeout = 600</c>.
    /// </summary>
    private const int CommandTimeoutSeconds = 600;

    private readonly ILogger<SpExecutor> _logger;

    public SpExecutor(ILogger<SpExecutor> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc />
    public async Task ExecuteAsync(
        IDbConnection connection,
        string spName,
        object? parameters = null,
        IDbTransaction? transaction = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            _logger.LogDebug("Ejecutando SP {SpName}", spName);

            var commandDefinition = new CommandDefinition(
                spName,
                parameters,
                transaction,
                CommandTimeoutSeconds,
                CommandType.StoredProcedure,
                cancellationToken: cancellationToken);

            await connection.ExecuteAsync(commandDefinition);
        }
        catch (SqlException ex)
        {
            throw SqlExceptionTranslator.Translate(spName, ex);
        }
    }

    /// <inheritdoc />
    public async Task<IEnumerable<T>> QueryAsync<T>(
        IDbConnection connection,
        string spName,
        object? parameters = null,
        IDbTransaction? transaction = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            _logger.LogDebug("Consultando SP {SpName}", spName);

            var commandDefinition = new CommandDefinition(
                spName,
                parameters,
                transaction,
                CommandTimeoutSeconds,
                CommandType.StoredProcedure,
                cancellationToken: cancellationToken);

            return await connection.QueryAsync<T>(commandDefinition);
        }
        catch (SqlException ex)
        {
            throw SqlExceptionTranslator.Translate(spName, ex);
        }
    }

    /// <inheritdoc />
    public async Task<T?> QueryFirstOrDefaultAsync<T>(
        IDbConnection connection,
        string spName,
        object? parameters = null,
        IDbTransaction? transaction = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            _logger.LogDebug("Consultando primer resultado SP {SpName}", spName);

            var commandDefinition = new CommandDefinition(
                spName,
                parameters,
                transaction,
                CommandTimeoutSeconds,
                CommandType.StoredProcedure,
                cancellationToken: cancellationToken);

            return await connection.QueryFirstOrDefaultAsync<T>(commandDefinition);
        }
        catch (SqlException ex)
        {
            throw SqlExceptionTranslator.Translate(spName, ex);
        }
    }

    /// <inheritdoc />
    public async Task<IEnumerable<IEnumerable<object>>> QueryMultipleAsync(
        IDbConnection connection,
        string spName,
        object? parameters = null,
        IDbTransaction? transaction = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            _logger.LogDebug("Consultando múltiples resultsets SP {SpName}", spName);

            var commandDefinition = new CommandDefinition(
                spName,
                parameters,
                transaction,
                CommandTimeoutSeconds,
                CommandType.StoredProcedure,
                cancellationToken: cancellationToken);

            using var multi = await connection.QueryMultipleAsync(commandDefinition);
            var results = new List<IEnumerable<object>>();

            while (!multi.IsConsumed)
            {
                results.Add(await multi.ReadAsync<object>());
            }

            return results;
        }
        catch (SqlException ex)
        {
            throw SqlExceptionTranslator.Translate(spName, ex);
        }
    }
}
