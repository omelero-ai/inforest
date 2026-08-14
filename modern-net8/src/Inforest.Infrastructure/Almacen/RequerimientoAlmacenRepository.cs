using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Almacen;

namespace Inforest.Infrastructure.Almacen;

/// <summary>
/// Repositorio Dapper de requerimientos del módulo ALMACEN.
/// Legacy: frmImportacionRequerimientos.frm, frmImportacionRequerimientoDetalle.frm.
/// Accede a la BD ALMACEN via IDbConnectionFactory con nombre "ALMACEN".
/// BR-IMPORT-001, BR-IMPORT-002, BR-IMPORT-003, BR-IMPORT-004.
/// </summary>
internal sealed class RequerimientoAlmacenRepository : IRequerimientoAlmacenRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public RequerimientoAlmacenRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    /// <inheritdoc/>
    public async Task<IReadOnlyList<RequerimientoAlmacen>> ObtenerPendientesAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("ALMACEN", ct);

        // Legacy: cmdProcesa_Click en frmImportacionRequerimientos.frm
        // Filtra por CodEstado='02' (aprobado) y T.lImportarPV=1 (área habilitada para POS)
        const string sql = """
            SELECT DISTINCT
                v.Rq,
                v.Fecha,
                ISNULL(v.Area, '') AS Area,
                ISNULL(v.Responsable, '') AS Responsable,
                ISNULL(v.CodEstado, '') AS CodEstado,
                ISNULL(v.tPedido, '') AS CodigoPedido
            FROM vRequerimiento v
            INNER JOIN TRUTAAREA T ON v.CodigoArea = T.tCodigoArea
            WHERE v.CodEstado = '02'
              AND v.Fecha >= @fechaInicio
              AND v.Fecha <= @fechaFin
              AND T.lImportarPV = 1
            ORDER BY v.Fecha DESC
            """;

        var rows = await conn.QueryAsync<RequerimientoRow>(sql, new
        {
            fechaInicio = fechaInicio.Date,
            fechaFin = fechaFin.Date.AddHours(23).AddMinutes(59)
        });

        return rows.Select(r => RequerimientoAlmacen.Crear(
            r.Rq,
            r.Fecha,
            r.Area,
            r.Responsable,
            r.CodEstado,
            string.IsNullOrWhiteSpace(r.CodigoPedido) ? null : r.CodigoPedido))
            .ToList();
    }

    /// <inheritdoc/>
    public async Task<IReadOnlyList<DetalleRequerimientoAlmacen>> ObtenerDetalleAsync(
        string rq,
        string baseDatosInforest,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("ALMACEN", ct);

        // Legacy: frmImportacionRequerimientoDetalle.frm Form_Load / Asignar
        // Hace LEFT JOIN a TPRODUCTO en la BD de INFOREST para obtener el enlace
        var sql = $"""
            SELECT
                V.CodArt,
                ISNULL(V.Articulo, '') AS Articulo,
                ISNULL(V.nAtendido, V.Cantidad) AS Cantidad,
                ISNULL(T.tCodigoProducto, '') AS EnlaceInforRest
            FROM vRequerimiento V
            LEFT JOIN [{baseDatosInforest}].dbo.TPRODUCTO T ON V.CodArt = T.tEnlace
            WHERE V.Rq = @rq
            """;

        var rows = await conn.QueryAsync<DetalleRequerimientoRow>(sql, new { rq });

        return rows.Select(r => DetalleRequerimientoAlmacen.Crear(
            r.CodArt,
            r.Articulo,
            r.Cantidad,
            string.IsNullOrWhiteSpace(r.EnlaceInforRest) ? null : r.EnlaceInforRest))
            .ToList();
    }

    /// <inheritdoc/>
    public async Task<string?> ObtenerPedidoAsociadoAsync(
        string rq,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("ALMACEN", ct);

        // Legacy: Calcular(" Select ISNULL(tPedido,'') As Codigo From MREQUERIMIENTO Where tRequerimiento = '...'", CnAlmacen)
        const string sql = """
            SELECT ISNULL(tPedido, '') AS Codigo
            FROM MREQUERIMIENTO
            WHERE tRequerimiento = @rq
            """;

        var codigo = await conn.QueryFirstOrDefaultAsync<string>(sql, new { rq });
        return string.IsNullOrWhiteSpace(codigo) ? null : codigo;
    }

    /// <inheritdoc/>
    public async Task MarcarImportadoAsync(
        string rq,
        string codigoPedido,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("ALMACEN", ct);

        // Legacy: "Update mRequerimiento Set lPedido = 1, tPedido = '...' Where tRequerimiento = '...'"
        const string sql = """
            UPDATE MREQUERIMIENTO
            SET lPedido = 1,
                tPedido = @codigoPedido
            WHERE tRequerimiento = @rq
            """;

        await conn.ExecuteAsync(sql, new { rq, codigoPedido });
    }

    // ── Projection records ───────────────────────────────────────────────────

    private sealed record RequerimientoRow(
        string Rq,
        DateTime Fecha,
        string Area,
        string Responsable,
        string CodEstado,
        string CodigoPedido);

    private sealed record DetalleRequerimientoRow(
        string CodArt,
        string Articulo,
        double Cantidad,
        string EnlaceInforRest);
}
