using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Precios;

namespace Inforest.Infrastructure.Precios;

/// <summary>
/// Repositorio de precios — vProducto + DPEDIDO + TLOG_MODPRECIO.
/// Legacy: frmPrecios.frm, frmPreciosDetalle.frm
/// BR-PRECIO-001..006
/// </summary>
internal sealed class PrecioRepository : IPrecioRepository
{
    private readonly IDbConnectionFactory _factory;

    public PrecioRepository(IDbConnectionFactory factory) => _factory = factory;

    /// <inheritdoc/>
    public async Task<IEnumerable<PrecioProductoFila>> ObtenerListadoPreciosAsync(CancellationToken ct = default)
    {
        // Legacy: frmPrecios.frm — SELECT Codigo, Grupo, SubGrupo, Descripcion,
        //         nPrecioVenta, nPrecioDelivery, nPrecioLlevar, lActivo FROM vProducto
        using var cn = await _factory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT
                Codigo,
                ISNULL(Grupo,'')      AS Grupo,
                ISNULL(SubGrupo,'')   AS SubGrupo,
                ISNULL(Descripcion,'') AS Descripcion,
                ISNULL(nPrecioVenta,0)    AS PrecioVenta,
                ISNULL(nPrecioDelivery,0) AS PrecioDelivery,
                ISNULL(nPrecioLlevar,0)   AS PrecioLlevar,
                ISNULL(CAST(lActivo AS BIT), 0) AS Activo
            FROM vProducto
            ORDER BY Grupo, SubGrupo, Descripcion
            """;

        var rows = await cn.QueryAsync<PrecioFilaRow>(new CommandDefinition(sql, cancellationToken: ct));
        return rows.Select(r => new PrecioProductoFila(
            r.Codigo, r.Grupo, r.SubGrupo, r.Descripcion,
            r.PrecioVenta, r.PrecioDelivery, r.PrecioLlevar, r.Activo));
    }

    /// <inheritdoc/>
    public async Task<ItemPedidoParaPrecio?> ObtenerItemParaPrecioAsync(
        string codigoPedido, string item, CancellationToken ct = default)
    {
        // Legacy: frmPreciosDetalle.frm — SELECT * FROM vPedidoDetalle
        //         WHERE tEstadoItem='N' AND tCodigoPedido=@p AND tItem=@i
        using var cn = await _factory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT
                d.tCodigoPedido                     AS CodigoPedido,
                d.tItem                             AS Item,
                d.tCodigoProducto                   AS CodigoProducto,
                ISNULL(p.tDetallado,'')             AS Producto,
                ISNULL(CAST(p.lModificable AS BIT),0) AS Modificable,
                ISNULL(d.nPrecioOficial,0)          AS PrecioOficial,
                ISNULL(d.nPrecioNeto,0)             AS PrecioNeto,
                ISNULL(d.nPrecioVenta,0)            AS PrecioVenta,
                ISNULL(d.nDescuento,0)              AS Descuento,
                ISNULL(d.nRecargo,0)                AS Recargo,
                ISNULL(d.nprecioImpuesto1,0)        AS Impuesto1,
                ISNULL(d.nprecioImpuesto2,0)        AS Impuesto2,
                ISNULL(d.nprecioImpuesto3,0)        AS Impuesto3,
                ISNULL(d.nCantidad,0)               AS Cantidad,
                ISNULL(CAST(d.tAplicaExoneracion AS BIT),0) AS AplicaExoneracion,
                ISNULL(CAST(d.tAplicaInafectacion AS BIT),0) AS AplicaInafectacion
            FROM DPEDIDO d
            LEFT JOIN TPRODUCTO p ON p.tCodigoProducto = d.tCodigoProducto
            WHERE d.tEstadoItem = 'N'
              AND d.tCodigoPedido = @codigoPedido
              AND d.tItem = @item
            """;

        var row = await cn.QueryFirstOrDefaultAsync<ItemPedidoFilaRow>(
            new CommandDefinition(sql, new { codigoPedido, item }, cancellationToken: ct));

        if (row is null) return null;

        return new ItemPedidoParaPrecio(
            row.CodigoPedido, row.Item, row.CodigoProducto, row.Producto,
            row.Modificable, row.PrecioOficial, row.PrecioNeto, row.PrecioVenta,
            row.Descuento, row.Recargo, row.Impuesto1, row.Impuesto2, row.Impuesto3,
            row.Cantidad, row.AplicaExoneracion, row.AplicaInafectacion);
    }

    /// <inheritdoc/>
    public async Task ModificarPrecioItemAsync(ModificarPrecioItemCommand cmd, CancellationToken ct = default)
    {
        // Legacy: frmPreciosDetalle.frm — GrabaProducto
        // UPDATE DPEDIDO + INSERT TLOG_MODPRECIO
        using var cn = await _factory.CreateOpenConnectionAsync("Inforest", ct);
        using var tx = cn.BeginTransaction();

        // BR-PRECIO-002: actualizar DPEDIDO
        const string sqlUpdate = """
            UPDATE DPEDIDO SET
                nPrecioNeto      = @precioNeto,
                nDescuento       = @descuento,
                nRecargo         = @recargo,
                nPrecioOficial   = @precioOficial,
                nprecioImpuesto1 = @impuesto1,
                nprecioImpuesto2 = @impuesto2,
                nprecioImpuesto3 = @impuesto3,
                nPrecioVenta     = @precioVenta,
                nventa           = @venta,
                nCantidad        = @cantidad,
                nImpuesto1       = @totalImp1,
                nImpuesto2       = @totalImp2,
                nImpuesto3       = @totalImp3
            WHERE tItem = @item
              AND tCodigoPedido = @codigoPedido
            """;

        await cn.ExecuteAsync(new CommandDefinition(sqlUpdate, new
        {
            precioNeto    = cmd.PrecioNeto,
            descuento     = cmd.Descuento,
            recargo       = cmd.Recargo,
            precioOficial = cmd.PrecioOficial,
            impuesto1     = cmd.Impuesto1,
            impuesto2     = cmd.Impuesto2,
            impuesto3     = cmd.Impuesto3,
            precioVenta   = cmd.PrecioVenta,
            venta         = cmd.PrecioVenta * cmd.Cantidad,
            cantidad      = cmd.Cantidad,
            totalImp1     = cmd.Impuesto1 * cmd.Cantidad,
            totalImp2     = cmd.Impuesto2 * cmd.Cantidad,
            totalImp3     = cmd.Impuesto3 * cmd.Cantidad,
            item          = cmd.Item,
            codigoPedido  = cmd.CodigoPedido
        }, transaction: tx, cancellationToken: ct));

        // BR-PRECIO-002: auditoría en TLOG_MODPRECIO
        const string sqlAudit = """
            INSERT INTO TLOG_MODPRECIO
                SELECT @codigoPedido, @item,
                    ndescuento,       @descuento,
                    nrecargo,         @recargo,
                    npreciooficial,   @precioOficial,
                    nprecioimpuesto1, @impuesto1,
                    nprecioimpuesto2, @impuesto2,
                    nprecioimpuesto3, @impuesto3,
                    nprecioventa,     @precioVenta,
                    nventa,           @venta,
                    ncantidad,        @cantidad,
                    nimpuesto1,       @totalImp1,
                    nimpuesto2,       @totalImp2,
                    nimpuesto3,       @totalImp3,
                    @usuarioAutoriza, @usuario, GETDATE()
                FROM DPEDIDO
                WHERE tItem = @item AND tCodigoPedido = @codigoPedido
            """;

        await cn.ExecuteAsync(new CommandDefinition(sqlAudit, new
        {
            codigoPedido    = cmd.CodigoPedido,
            item            = cmd.Item,
            descuento       = cmd.Descuento,
            recargo         = cmd.Recargo,
            precioOficial   = cmd.PrecioOficial,
            impuesto1       = cmd.Impuesto1,
            impuesto2       = cmd.Impuesto2,
            impuesto3       = cmd.Impuesto3,
            precioVenta     = cmd.PrecioVenta,
            venta           = cmd.PrecioVenta * cmd.Cantidad,
            cantidad        = cmd.Cantidad,
            totalImp1       = cmd.Impuesto1 * cmd.Cantidad,
            totalImp2       = cmd.Impuesto2 * cmd.Cantidad,
            totalImp3       = cmd.Impuesto3 * cmd.Cantidad,
            usuarioAutoriza = cmd.UsuarioAutoriza,
            usuario         = cmd.Usuario
        }, transaction: tx, cancellationToken: ct));

        tx.Commit();
    }

    // ── Private rows ──────────────────────────────────────────────────────────

    private sealed record PrecioFilaRow(
        string Codigo, string Grupo, string SubGrupo, string Descripcion,
        decimal PrecioVenta, decimal PrecioDelivery, decimal PrecioLlevar, bool Activo);

    private sealed record ItemPedidoFilaRow(
        string CodigoPedido, string Item, string CodigoProducto, string Producto,
        bool Modificable, decimal PrecioOficial, decimal PrecioNeto, decimal PrecioVenta,
        decimal Descuento, decimal Recargo, decimal Impuesto1, decimal Impuesto2, decimal Impuesto3,
        decimal Cantidad, bool AplicaExoneracion, bool AplicaInafectacion);
}
