namespace Inforest.Domain.Repositories;

/// <summary>
/// Unidad de trabajo para coordinar transacciones.
/// Legacy: Las transacciones en INFOREST se manejan en los SPs (BEGIN TRAN / COMMIT / ROLLBACK).
/// En .NET 8 se coordina desde Application layer.
/// </summary>
public interface IUnitOfWork
{
    Task<int> CommitAsync(CancellationToken cancellationToken = default);
    Task RollbackAsync(CancellationToken cancellationToken = default);
}
