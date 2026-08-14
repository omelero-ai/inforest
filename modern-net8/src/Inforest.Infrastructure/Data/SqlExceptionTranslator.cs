using Inforest.Infrastructure.Exceptions;
using Microsoft.Data.SqlClient;

namespace Inforest.Infrastructure.Data;

/// <summary>
/// Traduce <see cref="SqlException"/> a <see cref="InfrastructureException"/> con contexto del SP.
/// <para>
/// Legacy: <c>clsComando.cls</c> mostraba el error como <c>MsgBox err.Description</c>.
/// En .NET 8, los errores SQL se encapsulan como excepciones tipificadas para permitir
/// manejo diferenciado por capa (log, reintento, mensaje de usuario).
/// </para>
/// Errores SQL de negocio (RAISERROR con severidad ≤ 16) se distinguen de errores de
/// infraestructura (conexión perdida, timeout, etc.).
/// </summary>
internal static class SqlExceptionTranslator
{
    // Número de error SQL para timeout de query
    private const int SqlErrorQueryTimeout = -2;

    // Rango de severidades consideradas errores de negocio en el Legacy
    private const byte MaxBusinessSeverity = 16;

    /// <summary>
    /// Convierte una <see cref="SqlException"/> en <see cref="InfrastructureException"/> enriquecida.
    /// </summary>
    public static InfrastructureException Translate(string spName, SqlException ex)
    {
        if (ex.Number == SqlErrorQueryTimeout)
        {
            return new InfrastructureException(
                $"Timeout al ejecutar SP '{spName}' (límite: 600 s). " +
                "Verificar carga del servidor o revisar el plan de ejecución del SP.",
                ex);
        }

        if (ex.Class <= MaxBusinessSeverity)
        {
            // Error lanzado desde el SP con RAISERROR — contiene mensaje de negocio
            return new InfrastructureException(
                $"El SP '{spName}' retornó un error de negocio: {ex.Message}",
                ex);
        }

        // Error de infraestructura (conexión perdida, BD inaccesible, etc.)
        return new InfrastructureException(
            $"Error de infraestructura SQL al ejecutar SP '{spName}': [{ex.Number}] {ex.Message}",
            ex);
    }
}
