using Dapper;
using Inforest.Application.Caja;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Hardware;

namespace Inforest.Infrastructure.Hardware;

internal sealed class PinPadTerminalRepository : IPinPadTerminalRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public PinPadTerminalRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<TerminalPinPad>> ObtenerActivosPorCajaAsync(string codigoCaja, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                CT.tTerminal AS CodigoTerminal,
                COALESCE(NULLIF(T.tDescripcion, ''), T.tTerminal, CT.tTerminal) AS Descripcion,
                ISNULL(T.tIntegracion, '') AS CodigoIntegracion
            FROM TCAJATERMINAL CT
            INNER JOIN TTERMINAL T ON CT.tTerminal = T.Codigo
            WHERE CT.tCaja = @CodigoCaja
              AND ISNULL(CT.lActivo, 0) = 1
              AND T.tIntegracion = '01'
            ORDER BY T.Codigo
            """;

        var rows = await connection.QueryAsync<TerminalRow>(
            new CommandDefinition(sql, new { CodigoCaja = codigoCaja.Trim() }, cancellationToken: ct));

        return rows
            .Select(row => new TerminalPinPad(
                row.CodigoTerminal,
                row.Descripcion,
                row.CodigoIntegracion))
            .ToList();
    }

    private sealed record TerminalRow(string CodigoTerminal, string Descripcion, string CodigoIntegracion);
}
