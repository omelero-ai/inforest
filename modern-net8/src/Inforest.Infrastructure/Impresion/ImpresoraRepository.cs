using Dapper;
using Inforest.Application.Impresion;
using Inforest.Application.Interfaces;

namespace Inforest.Infrastructure.Impresion;

/// <summary>
/// Repositorio Dapper para impresoras disponibles.
/// Legacy: TIMPRESORA. frmPrecuentaImpresora.frm Form_Load.
/// </summary>
internal sealed class ImpresoraRepository : IImpresoraRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public ImpresoraRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<ImpresoraDisponible>> ObtenerImpresorasPorCajaAsync(
        string codigoCaja,
        CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tImpresora AS CodigoImpresora,
                   tDescripcion AS Descripcion
              FROM TIMPRESORA
             WHERE tCaja = @CodigoCaja
             ORDER BY tImpresora
            """;

        var rows = await connection.QueryAsync<ImpresoraDisponible>(
            new CommandDefinition(sql, new { CodigoCaja = codigoCaja }, cancellationToken: ct));

        return rows.ToList().AsReadOnly();
    }
}
