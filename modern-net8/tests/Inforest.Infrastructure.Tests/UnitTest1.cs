using Inforest.Infrastructure.Exceptions;

namespace Inforest.Infrastructure.Tests;

/// <summary>
/// Pruebas baseline de P3-01 — valida contratos de excepción de Infrastructure.
/// </summary>
public class InfrastructureExceptionTests
{
    [Fact]
    public void InfrastructureException_TieneMensajeYCodigo()
    {
        var ex = new InfrastructureException("fallo técnico", "INFRA_ERR");
        Assert.Equal("fallo técnico", ex.Message);
        Assert.Equal("INFRA_ERR", ex.CodigoError);
    }

    [Fact]
    public void DatabaseException_CodigoEsDB_ERROR()
    {
        var inner = new Exception("sql error");
        var ex = new DatabaseException("error de base de datos", inner);
        Assert.Equal("DB_ERROR", ex.CodigoError);
        Assert.Same(inner, ex.InnerException);
    }

    [Fact]
    public void DatabaseException_EsInfrastructureException()
    {
        var ex = new DatabaseException("msg");
        Assert.IsAssignableFrom<InfrastructureException>(ex);
    }
}