using System.Data;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Ejecutor centralizado de stored procedures Legacy.
/// Abstrae los cuatro patrones de uso que expone <c>clsComando.cls</c>:
/// <list type="bullet">
///   <item>ExecSP  — ejecución sin resultado</item>
///   <item>GetSP   — lectura de un recordset</item>
///   <item>GetSP (primer registro) — lectura de un único objeto</item>
///   <item>GetParameterValue — lectura de parámetros OUTPUT</item>
/// </list>
/// Todos los SPs se invocan con el nombre exacto del Legacy (p. ej. <c>spIns_MPEDIDO</c>).
/// El timeout se hereda del Legacy: 600 segundos (<c>Cmd.CommandTimeout = 600</c>).
/// </summary>
public interface ISpExecutor
{
    /// <summary>
    /// Ejecuta un SP sin retornar filas (equivale a <c>ExecSP()</c> en <c>clsComando.cls</c>).
    /// </summary>
    Task ExecuteAsync(
        IDbConnection connection,
        string spName,
        object? parameters = null,
        IDbTransaction? transaction = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Ejecuta un SP y retorna una colección de objetos mapeados.
    /// Equivale a <c>GetSP()</c> + iteración del recordset.
    /// </summary>
    Task<IEnumerable<T>> QueryAsync<T>(
        IDbConnection connection,
        string spName,
        object? parameters = null,
        IDbTransaction? transaction = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Ejecuta un SP y retorna el primer objeto mapeado o <c>null</c>.
    /// Equivale a <c>GetSP()</c> con acceso al primer registro del recordset.
    /// </summary>
    Task<T?> QueryFirstOrDefaultAsync<T>(
        IDbConnection connection,
        string spName,
        object? parameters = null,
        IDbTransaction? transaction = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Ejecuta un SP y retorna múltiples resultsets.
    /// Legacy: algunos SPs retornan más de un SELECT; este método los expone mediante
    /// <see cref="SqlMapper.GridReader"/> de Dapper.
    /// </summary>
    Task<IEnumerable<IEnumerable<object>>> QueryMultipleAsync(
        IDbConnection connection,
        string spName,
        object? parameters = null,
        IDbTransaction? transaction = null,
        CancellationToken cancellationToken = default);
}
