using System.Data;
using Inforest.Infrastructure.Data;
using Inforest.Infrastructure.Exceptions;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace Inforest.Infrastructure.Tests.Data;

/// <summary>
/// Pruebas unitarias para <see cref="UnitOfWork"/>.
/// </summary>
public sealed class UnitOfWorkTests
{
    private static (UnitOfWork uow, Mock<IDbConnection> connMock, Mock<IDbTransaction> txMock)
        BuildUow()
    {
        var txMock = new Mock<IDbTransaction>();
        var connMock = new Mock<IDbConnection>();
        connMock.Setup(c => c.BeginTransaction(It.IsAny<IsolationLevel>())).Returns(txMock.Object);

        var uow = new UnitOfWork(connMock.Object, NullLogger<UnitOfWork>.Instance);
        return (uow, connMock, txMock);
    }

    [Fact]
    public void BeginTransaction_WithoutPrior_SetsTransaction()
    {
        var (uow, _, _) = BuildUow();

        uow.BeginTransaction();

        Assert.NotNull(uow.Transaction);
    }

    [Fact]
    public void BeginTransaction_TwiceThrows()
    {
        var (uow, _, _) = BuildUow();
        uow.BeginTransaction();

        Assert.Throws<InfrastructureException>(() => uow.BeginTransaction());
    }

    [Fact]
    public async Task CommitAsync_WithoutTransaction_Throws()
    {
        var (uow, _, _) = BuildUow();

        await Assert.ThrowsAsync<InfrastructureException>(() => uow.CommitAsync());
    }

    [Fact]
    public async Task CommitAsync_WithTransaction_CallsCommit()
    {
        var (uow, _, txMock) = BuildUow();
        uow.BeginTransaction();

        await uow.CommitAsync();

        txMock.Verify(t => t.Commit(), Times.Once);
    }

    [Fact]
    public async Task RollbackAsync_WithoutTransaction_DoesNotThrow()
    {
        var (uow, _, _) = BuildUow();

        // No debe lanzar
        await uow.RollbackAsync();
    }

    [Fact]
    public async Task RollbackAsync_WithTransaction_CallsRollback()
    {
        var (uow, _, txMock) = BuildUow();
        uow.BeginTransaction();

        await uow.RollbackAsync();

        txMock.Verify(t => t.Rollback(), Times.Once);
    }

    [Fact]
    public async Task DisposeAsync_WithActiveTransaction_ImplicitRollback()
    {
        var (uow, _, txMock) = BuildUow();
        uow.BeginTransaction();

        await uow.DisposeAsync();

        txMock.Verify(t => t.Rollback(), Times.Once);
    }
}
