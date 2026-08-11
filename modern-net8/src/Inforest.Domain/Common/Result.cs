namespace Inforest.Domain.Common;

/// <summary>
/// Contrato base de resultado para comandos y queries de Application.
/// Evita excepciones como control de flujo en casos de negocio esperados.
///
/// Uso: devolver Result.Ok() o Result.Fail("mensaje") desde servicios y handlers.
/// Convención INFOREST: siempre anotar el código de error con un prefijo de módulo,
/// p.ej. "PEDIDO_TURNO_REQUERIDO", "SEGURIDAD_USUARIO_INACTIVO".
/// </summary>
public sealed class Result
{
    public bool EsExitoso { get; }
    public string? CodigoError { get; }
    public string? MensajeError { get; }

    private Result(bool esExitoso, string? codigoError, string? mensajeError)
    {
        EsExitoso = esExitoso;
        CodigoError = codigoError;
        MensajeError = mensajeError;
    }

    public static Result Ok() => new(true, null, null);

    public static Result Fail(string mensaje, string? codigoError = null)
        => new(false, codigoError, mensaje);

    public static Result<T> Ok<T>(T valor) => Result<T>.Ok(valor);

    public static Result<T> Fail<T>(string mensaje, string? codigoError = null)
        => Result<T>.Fail(mensaje, codigoError);
}

/// <summary>
/// Contrato de resultado tipado para queries que retornan datos.
/// </summary>
public sealed class Result<T>
{
    public bool EsExitoso { get; }
    public T? Valor { get; }
    public string? CodigoError { get; }
    public string? MensajeError { get; }

    private Result(bool esExitoso, T? valor, string? codigoError, string? mensajeError)
    {
        EsExitoso = esExitoso;
        Valor = valor;
        CodigoError = codigoError;
        MensajeError = mensajeError;
    }

    public static Result<T> Ok(T valor) => new(true, valor, null, null);

    public static Result<T> Fail(string mensaje, string? codigoError = null)
        => new(false, default, codigoError, mensaje);

    public static implicit operator Result(Result<T> resultado)
        => resultado.EsExitoso
            ? Result.Ok()
            : Result.Fail(resultado.MensajeError!, resultado.CodigoError);
}
