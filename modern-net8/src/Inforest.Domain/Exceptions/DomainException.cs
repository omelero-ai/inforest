namespace Inforest.Domain.Exceptions;

/// <summary>
/// Excepción base para reglas de negocio del dominio INFOREST.
/// Las reglas de negocio violadas deben lanzar esta excepción o derivadas.
/// </summary>
public class DomainException : Exception
{
    public string? Code { get; }

    public DomainException(string message, string? code = null)
        : base(message)
    {
        Code = code;
    }

    public DomainException(string message, Exception innerException, string? code = null)
        : base(message, innerException)
    {
        Code = code;
    }
}
