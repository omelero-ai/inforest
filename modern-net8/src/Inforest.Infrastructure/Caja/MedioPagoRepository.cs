using Dapper;
using Inforest.Application.Caja;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Caja;

namespace Inforest.Infrastructure.Caja;

/// <summary>
/// Repositorio Dapper de catálogo de medios de pago.
/// Legacy asumido por workstream: TMEDIODEPAGO.
/// </summary>
internal sealed class MedioPagoRepository : IMedioPagoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public MedioPagoRepository(IDbConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory;
    }

    public async Task<IReadOnlyList<MedioPago>> ObtenerActivosAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                tMedioPago AS Codigo,
                tDetallado AS Descripcion,
                tTipo AS Tipo,
                lActivo AS Activo
            FROM TMEDIODEPAGO
            WHERE lActivo = 1
            ORDER BY tDetallado
            """;

        var rows = await connection.QueryAsync<MedioPagoRow>(
            new CommandDefinition(sql, cancellationToken: ct));

        return rows
            .Select(r => MedioPago.Reconstituir(
                r.Codigo,
                r.Descripcion,
                MapTipo(r.Tipo),
                r.Activo))
            .ToList();
    }

    private static TipoMedioPago MapTipo(string? tipo)
        => (tipo ?? string.Empty).Trim().ToUpperInvariant() switch
        {
            "EFE" => TipoMedioPago.Efectivo,
            "TAR" => TipoMedioPago.Tarjeta,
            "CHE" => TipoMedioPago.Cheque,
            "VAL" => TipoMedioPago.Vale,
            _ => TipoMedioPago.Otro
        };

    private sealed record MedioPagoRow(string Codigo, string Descripcion, string? Tipo, bool Activo);
}
