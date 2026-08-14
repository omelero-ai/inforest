using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio Dapper de insumos/platos de stock crítico (TINSUMO).
/// Legacy: frmInsumo.frm, frmInsumoDetalle.frm, clsAlmacen.cls.
/// SPs: USP_LISTARINSUMOS, usp_agregarinsumos, USP_MODIFICARINSUMOS, USP_ELIMINARINSUMOS.
/// BR-INSUMO-001, BR-INSUMO-002, BR-INSUMO-003, BR-INSUMO-004.
/// </summary>
internal sealed class InsumoRepository : IInsumoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public InsumoRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    /// <inheritdoc/>
    public async Task<IReadOnlyList<Insumo>> ObtenerTodosAsync(
        bool moduloInforRest = true, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        var rows = await _spExecutor.QueryAsync<InsumoRow>(
            conn,
            "USP_LISTARINSUMOS",
            new { modulo = moduloInforRest });
        return rows.Select(MapRow).ToList();
    }

    /// <inheritdoc/>
    public async Task<Insumo?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigo AS Codigo,
                   tusuarioReg AS Usuario,
                   DESCRIPCION AS Descripcion,
                   ISNULL(nStock, 0) AS Stock,
                   fRegistro AS FechaRegistro,
                   lActivo AS Activo,
                   ISNULL(lINSUMO, 1) AS EsInsumo
            FROM TINSUMO
            WHERE tCodigo = @codigo
            """;
        var row = await conn.QueryFirstOrDefaultAsync<InsumoRow>(sql, new { codigo });
        return row is null ? null : MapRow(row);
    }

    /// <inheritdoc/>
    public async Task<bool> ExisteDescripcionAsync(string descripcion, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT COUNT(*) FROM TINSUMO
            WHERE Descripcion = UPPER(@descripcion)
            """;
        var count = await conn.ExecuteScalarAsync<int>(sql, new { descripcion });
        return count > 0;
    }

    /// <inheritdoc/>
    public async Task<string> ObtenerProximoCodigoAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = "SELECT MAX(tCodigo) FROM TINSUMO";
        var maxCodigo = await conn.ExecuteScalarAsync<string?>(sql);

        if (string.IsNullOrEmpty(maxCodigo))
            return "00000001";

        // BR-INSUMO-002: correlativo numérico de 8 dígitos
        if (long.TryParse(maxCodigo, out var numero))
            return (numero + 1).ToString("D8");

        return "00000001";
    }

    /// <inheritdoc/>
    public async Task AgregarAsync(
        Insumo insumo, string caja, bool moduloInforRest = true, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(
            conn,
            "usp_agregarinsumos",
            new
            {
                codigo      = insumo.Codigo,
                usuario     = insumo.Usuario,
                descripcion = insumo.Descripcion,
                nstock      = insumo.Stock,
                tCaja       = caja,
                activo      = insumo.Activo,
                LINSUMO     = insumo.EsInsumo
            });
    }

    /// <inheritdoc/>
    public async Task ModificarAsync(
        Insumo insumo, string caja, bool moduloInforRest = true, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(
            conn,
            "USP_MODIFICARINSUMOS",
            new
            {
                codigo      = insumo.Codigo,
                usuario     = insumo.Usuario,
                descripcion = insumo.Descripcion,
                tCaja       = caja,
                activo      = insumo.Activo,
                nstock      = insumo.Stock,
                MODULO      = moduloInforRest,
                LINSUMO     = insumo.EsInsumo
            });
    }

    /// <inheritdoc/>
    public async Task EliminarAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(
            conn,
            "USP_ELIMINARINSUMOS",
            new { codigo });
    }

    private static Insumo MapRow(InsumoRow r) =>
        Insumo.Crear(
            r.Codigo,
            r.Descripcion,
            r.Usuario,
            r.Stock,
            r.EsInsumo,
            r.Activo,
            r.FechaRegistro);

    private sealed record InsumoRow(
        string Codigo,
        string Usuario,
        string Descripcion,
        double Stock,
        DateTime FechaRegistro,
        bool Activo,
        bool EsInsumo);
}
