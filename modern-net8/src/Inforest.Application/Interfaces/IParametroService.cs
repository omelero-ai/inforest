namespace Inforest.Application.Interfaces;

/// <summary>
/// Abstracción para acceso a configuración del sistema.
/// Legacy: Tabla TPARAMETRO — tabla maestra de configuración con 100+ flags.
/// Los valores de TPARAMETRO determinan comportamiento de muchas funcionalidades.
/// </summary>
public interface IParametroService
{
    /// <summary>
    /// Obtiene el valor de un parámetro del sistema por su clave.
    /// Legacy: SELECT tValor FROM TPARAMETRO WHERE tParametro = @clave
    /// </summary>
    Task<string?> ObtenerAsync(string clave, CancellationToken cancellationToken = default);

    /// <summary>
    /// Obtiene el valor numérico de un parámetro.
    /// </summary>
    Task<decimal?> ObtenerNumericoAsync(string clave, CancellationToken cancellationToken = default);

    /// <summary>
    /// Obtiene el valor booleano de un parámetro.
    /// Legacy: "SI"/"NO" o "1"/"0" en TPARAMETRO
    /// </summary>
    Task<bool> ObtenerBoolAsync(string clave, bool valorDefecto = false, CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifica si existe configuración global cargable desde TPARAMETRO.
    /// </summary>
    Task<bool> TieneConfiguracionAsync(CancellationToken cancellationToken = default);

    void InvalidarCache();
}
