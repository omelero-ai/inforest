using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Infrastructure.Ventas;

/// <summary>
/// Repositorio de motivos de anulación sobre SQL Server.
/// Legacy: vMotivoAnulacion — SELECT SUBSTRING(TCODIGO,1,3) AS Codigo, tDetallado AS Descripcion,
///         tResumido, lActivo, nValor FROM TTABLA WHERE TTABLA = N'MOTIVOANULACION'.
/// BR-MOTIVO-001.
/// </summary>
public sealed class MotivoAnulacionRepository(IDbConnectionFactory connectionFactory)
    : IMotivoAnulacionRepository
{
    private const string Sql =
        """
        SELECT SUBSTRING(TCODIGO, 1, 3) AS Codigo,
               tDetallado               AS Descripcion,
               tResumido,
               CAST(lActivo AS BIT)     AS Activo,
               nValor                  AS ValorNumerico
        FROM   dbo.TTABLA
        WHERE  TTABLA = N'MOTIVOANULACION'
          AND  ISNULL(lActivo, 0) = 1
        ORDER  BY TCODIGO
        """;

    /// <inheritdoc/>
    public async Task<IReadOnlyList<MotivoAnulacion>> ObtenerActivosAsync(CancellationToken ct = default)
    {
        using var cn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var rows = await cn.QueryAsync<MotivoAnulacionRow>(
            new CommandDefinition(Sql, cancellationToken: ct));

        return rows.Select(r => new MotivoAnulacion(
            r.Codigo,
            r.Descripcion,
            r.TResumido ?? string.Empty,
            r.Activo,
            r.ValorNumerico))
            .ToList()
            .AsReadOnly();
    }

    private sealed record MotivoAnulacionRow(
        string Codigo,
        string Descripcion,
        string? TResumido,
        bool Activo,
        double? ValorNumerico);
}
