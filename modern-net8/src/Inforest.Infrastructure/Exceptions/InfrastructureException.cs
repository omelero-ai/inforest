namespace Inforest.Infrastructure.Exceptions;

/// <summary>
/// Excepción lanzada por la capa Infrastructure ante fallos técnicos no de dominio.
/// Ejemplos: error de conexión a SQL Server, timeout de SP, fallo de hardware.
///
/// Esta excepción no debe propagarse hacia Application ni Domain sin traducción.
/// Los handlers de Application deben convertirla a Result.Fail() con código de error apropiado.
/// </summary>
public class InfrastructureException : Exception
{
    public string? CodigoError { get; }

    public InfrastructureException(string message, string? codigoError = null)
        : base(message)
    {
        CodigoError = codigoError;
    }

    public InfrastructureException(string message, Exception innerException, string? codigoError = null)
        : base(message, innerException)
    {
        CodigoError = codigoError;
    }
}

/// <summary>
/// Error de acceso a base de datos SQL Server.
/// Legacy: errores de ADODB en clsComando.cls — On Error GoTo / Err.Number.
/// </summary>
public class DatabaseException : InfrastructureException
{
    public DatabaseException(string message, Exception? innerException = null)
        : base(message, innerException!, "DB_ERROR") { }
}
