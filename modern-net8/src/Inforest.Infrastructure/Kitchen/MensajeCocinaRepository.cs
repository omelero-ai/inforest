using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Kitchen;
using Inforest.Domain.Entities.Cocina;
using System.Globalization;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Repositorio Dapper de mensajes de cocina/KDS.
/// Legacy: frmMensajeCocina.frm, frmMensajeCocinaDetalle.frm, frmLiquidacionDetalle.frm.
/// SPs: USP_LISTARMENSAJES, USP_AGREGARMENSAJE, USP_MODIFICARMENSAJE, USP_ELIMINARRMENSAJES, USP_CERRAR_MENSAJES_CIERRETURNO.
/// Reglas BR-MSGCOC-001, BR-MSGCOC-002, BR-MSGCOC-003, BR-MSGCOC-004, BR-MSGCOC-005.
/// </summary>
internal sealed class MensajeCocinaRepository : IMensajeCocinaRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public MensajeCocinaRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<IReadOnlyList<MensajeCocina>> ObtenerMensajesAsync(DateTime fechaInicio, DateTime fechaFin, string codigoCaja, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        var rows = await _spExecutor.QueryAsync<MensajeCocinaListRow>(
            connection,
            "USP_LISTARMENSAJES",
            new
            {
                fechaini = fechaInicio.Date,
                fechafin = fechaFin.Date.AddDays(1).AddSeconds(-1),
                tcaja = codigoCaja ?? string.Empty
            },
            cancellationToken: ct);

        return rows.Select(MapListRow).ToList();
    }

    public async Task<MensajeCocina?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT Codigo,
                   tUsuarioReg AS UsuarioRegistro,
                   Mensaje,
                   fRegistro AS FechaRegistro,
                   fFinal AS FechaModificacion,
                   tUsuarioFinal AS UsuarioModificacion,
                   tCaja AS Caja,
                   lActivo AS Activo,
                   lActivarAlerta AS ActivarAlerta
            FROM TMENSAJECOCINA
            WHERE Codigo = @codigo
            """;

        var row = await connection.QueryFirstOrDefaultAsync<MensajeCocinaDetailRow>(new CommandDefinition(sql, new { codigo }, cancellationToken: ct));
        return row is null ? null : MapDetailRow(row);
    }

    public async Task<int> ContarActivosAsync(string? codigoExcluir = null, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT COUNT(Codigo)
            FROM TMENSAJECOCINA
            WHERE lActivo = 1
              AND (@codigoExcluir = '' OR Codigo <> @codigoExcluir)
            """;

        return await connection.ExecuteScalarAsync<int>(new CommandDefinition(sql, new { codigoExcluir = codigoExcluir ?? string.Empty }, cancellationToken: ct));
    }

    public async Task<string> ObtenerProximoCodigoAsync(DateTime fecha, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT MAX(Codigo)
            FROM TMENSAJECOCINA
            WHERE SUBSTRING(Codigo, 1, 2) = @prefijo
            """;

        var prefijo = fecha.ToString("yy", CultureInfo.InvariantCulture);
        var actual = await connection.ExecuteScalarAsync<string?>(new CommandDefinition(sql, new { prefijo }, cancellationToken: ct));
        if (string.IsNullOrWhiteSpace(actual) || actual.Length < 8 || !int.TryParse(actual[2..], out var correlativo))
            return prefijo + "000001";

        return prefijo + (correlativo + 1).ToString("D6", CultureInfo.InvariantCulture);
    }

    public async Task AgregarAsync(MensajeCocina mensaje, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(
            connection,
            "USP_AGREGARMENSAJE",
            new
            {
                codigo = mensaje.Codigo,
                usuario = mensaje.UsuarioRegistro,
                mensaje = mensaje.Mensaje,
                tcaja = mensaje.Caja,
                lactivo = mensaje.Activo,
                lactivaralerta = mensaje.ActivarAlerta
            },
            cancellationToken: ct);
    }

    public async Task ModificarAsync(MensajeCocina mensaje, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(
            connection,
            "USP_MODIFICARMENSAJE",
            new
            {
                usuario = mensaje.UsuarioModificacion,
                codigo = mensaje.Codigo,
                mensaje = mensaje.Mensaje,
                tcaja = mensaje.Caja,
                lactivo = mensaje.Activo,
                lactivaralerta = mensaje.ActivarAlerta
            },
            cancellationToken: ct);
    }

    public async Task EliminarAsync(string codigo, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(
            connection,
            "USP_ELIMINARRMENSAJES",
            new { tCODIGO = codigo },
            cancellationToken: ct);
    }

    public async Task CerrarActivosPorCajaAsync(string usuario, string codigoCaja, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(
            connection,
            "USP_CERRAR_MENSAJES_CIERRETURNO",
            new { usuario, tcaja = codigoCaja },
            cancellationToken: ct);
    }

    private static MensajeCocina MapListRow(MensajeCocinaListRow row)
        => MensajeCocina.Crear(
            row.Codigo ?? string.Empty,
            row.tusuarioreg ?? string.Empty,
            row.Mensaje ?? string.Empty,
            row.tCaja ?? string.Empty,
            row.lactivo,
            row.lActivarAlerta,
            ParseDate(row.fRegistro),
            ParseDate(row.ffinal),
            row.tusuariofinal);

    private static MensajeCocina MapDetailRow(MensajeCocinaDetailRow row)
        => MensajeCocina.Crear(
            row.Codigo,
            row.UsuarioRegistro,
            row.Mensaje,
            row.Caja,
            row.Activo,
            row.ActivarAlerta,
            row.FechaRegistro,
            row.FechaModificacion,
            row.UsuarioModificacion);

    private static DateTime ParseDate(string? value)
        => DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result)
            ? result
            : DateTime.MinValue;

    internal sealed class MensajeCocinaListRow
    {
        public string? Codigo { get; init; }
        public string? tusuarioreg { get; init; }
        public string? Mensaje { get; init; }
        public string? fRegistro { get; init; }
        public string? ffinal { get; init; }
        public string? tusuariofinal { get; init; }
        public string? tCaja { get; init; }
        public bool lactivo { get; init; }
        public bool lActivarAlerta { get; init; }
    }

    private sealed record MensajeCocinaDetailRow(
        string Codigo,
        string UsuarioRegistro,
        string Mensaje,
        DateTime FechaRegistro,
        DateTime FechaModificacion,
        string UsuarioModificacion,
        string Caja,
        bool Activo,
        bool ActivarAlerta);
}
