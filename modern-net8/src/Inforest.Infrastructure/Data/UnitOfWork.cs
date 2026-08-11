using System.Data;
using Inforest.Domain.Repositories;
using Inforest.Infrastructure.Exceptions;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Data;

/// <summary>
/// Implementación de <see cref="IUnitOfWork"/> que coordina transacciones ADO.NET.
/// <para>
/// Legacy: la mayoría de las transacciones en INFOREST ocurren dentro de los SPs
/// (BEGIN TRAN / COMMIT / ROLLBACK). Esta implementación extiende ese modelo al
/// permitir envolver múltiples operaciones .NET en una sola transacción de BD cuando
/// la lógica de negocio así lo requiera (p. ej. operaciones que combinan SPs con
/// escrituras directas a tabla).
/// </para>
/// Ciclo de vida: registrado como <c>Scoped</c>; cada request/operación obtiene su propia instancia.
/// </summary>
public sealed class UnitOfWork : IUnitOfWork, IAsyncDisposable
{
    private readonly IDbConnection _connection;
    private IDbTransaction? _transaction;
    private readonly ILogger<UnitOfWork> _logger;
    private bool _disposed;

    public UnitOfWork(IDbConnection connection, ILogger<UnitOfWork> logger)
    {
        _connection = connection;
        _logger = logger;
    }

    /// <summary>
    /// Conexión activa de esta unidad de trabajo.
    /// Disponible para repositorios que necesitan participar en la misma transacción.
    /// </summary>
    public IDbConnection Connection => _connection;

    /// <summary>
    /// Transacción activa; <c>null</c> si no se ha iniciado ninguna.
    /// </summary>
    public IDbTransaction? Transaction => _transaction;

    /// <summary>
    /// Inicia una nueva transacción en la conexión subyacente.
    /// </summary>
    public void BeginTransaction(IsolationLevel isolationLevel = IsolationLevel.ReadCommitted)
    {
        if (_transaction is not null)
            throw new InfrastructureException("Ya existe una transacción activa en esta unidad de trabajo.");

        _transaction = _connection.BeginTransaction(isolationLevel);
        _logger.LogDebug("Transacción iniciada con nivel de aislamiento {IsolationLevel}", isolationLevel);
    }

    /// <inheritdoc />
    public Task<int> CommitAsync(CancellationToken cancellationToken = default)
    {
        if (_transaction is null)
            throw new InfrastructureException("No hay transacción activa para confirmar.");

        try
        {
            _transaction.Commit();
            _logger.LogDebug("Transacción confirmada (COMMIT)");
            return Task.FromResult(0);
        }
        catch (SqlException ex)
        {
            throw new InfrastructureException("Error al confirmar la transacción.", ex);
        }
        finally
        {
            DisposeTransaction();
        }
    }

    /// <inheritdoc />
    public Task RollbackAsync(CancellationToken cancellationToken = default)
    {
        if (_transaction is null)
            return Task.CompletedTask;

        try
        {
            _transaction.Rollback();
            _logger.LogDebug("Transacción revertida (ROLLBACK)");
        }
        catch (SqlException ex)
        {
            _logger.LogError(ex, "Error al revertir la transacción");
        }
        finally
        {
            DisposeTransaction();
        }

        return Task.CompletedTask;
    }

    private void DisposeTransaction()
    {
        _transaction?.Dispose();
        _transaction = null;
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        _disposed = true;

        if (_transaction is not null)
        {
            _logger.LogWarning("UnitOfWork siendo descartado con transacción activa — haciendo ROLLBACK implícito");
            await RollbackAsync();
        }

        if (_connection is IAsyncDisposable asyncDisposable)
            await asyncDisposable.DisposeAsync();
        else
            _connection.Dispose();
    }
}
