using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Infrastructure.Ventas;

/// <summary>
/// Repositorio Dapper para documentos de venta.
/// Legacy: MDOCUMENTO / DDOCUMENTO. BR-002, BR-013.
/// </summary>
internal sealed class DocumentoRepository : IDocumentoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public DocumentoRepository(IDbConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory;
    }

    public async Task<Documento?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                tDocumento AS CodigoDocumento,
                tTipoDocumento AS TipoDocumento,
                tCodigoCliente AS CodigoCliente,
                nNeto AS Neto,
                nPrecioImpuesto1 AS Impuesto1,
                nPrecioImpuesto2 AS Impuesto2,
                nPrecioImpuesto3 AS Impuesto3,
                nVenta AS Venta,
                nDescuento AS Descuento,
                nPropina AS Propina,
                nTotal AS Total,
                nAbono AS Abono,
                nVuelto AS Vuelto,
                tEstadoDocumento AS Estado,
                tTurno AS CodigoTurno,
                tCaja AS CodigoCaja,
                fPago AS FechaPago,
                fRegistro AS FechaRegistro,
                fDiaContable AS FechaDiaContable,
                tUsuarioAnulado AS UsuarioAnulado,
                tObservacion AS MotivoAnulacion,
                fRegistroAnulado AS FechaAnulacion
            FROM MDOCUMENTO
            WHERE tDocumento = @Codigo
            """;

        var row = await connection.QueryFirstOrDefaultAsync<DocumentoRow>(
            new CommandDefinition(sql, new { Codigo = codigo }, cancellationToken: ct));

        return row is null
            ? null
            : Documento.Reconstituir(
                row.CodigoDocumento,
                row.TipoDocumento,
                row.CodigoCliente,
                row.Neto,
                row.Impuesto1,
                row.Impuesto2,
                row.Impuesto3,
                row.Venta,
                row.Descuento,
                row.Propina,
                row.Total,
                row.Estado,
                row.CodigoTurno,
                row.CodigoCaja,
                row.FechaRegistro,
                row.FechaPago,
                row.FechaDiaContable,
                row.Abono,
                row.Vuelto,
                row.UsuarioAnulado,
                row.MotivoAnulacion,
                row.FechaAnulacion);
    }

    public async Task<bool> InsertarAsync(Documento doc, IReadOnlyList<DetalleDocumento> detalles, CancellationToken ct = default)
    {
        if (detalles.Count == 0)
            return false;

        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        using var transaction = connection.BeginTransaction();

        try
        {
            const string insertDocumento = """
                INSERT INTO MDOCUMENTO
                (
                    tDocumento, tTipoDocumento, tCodigoCliente, nNeto,
                    nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3,
                    nVenta, nDescuento, nPropina, nTotal, nAbono, nVuelto,
                    tEstadoDocumento, tTurno, tCaja, fRegistro, fDiaContable
                )
                VALUES
                (
                    @CodigoDocumento, @TipoDocumento, @CodigoCliente, @Neto,
                    @Impuesto1, @Impuesto2, @Impuesto3,
                    @Venta, @Descuento, @Propina, @Total, @Abono, @Vuelto,
                    @Estado, @CodigoTurno, @CodigoCaja, @FechaRegistro, @FechaDiaContable
                )
                """;

            var cabecera = await connection.ExecuteAsync(new CommandDefinition(
                insertDocumento,
                new
                {
                    doc.CodigoDocumento,
                    doc.TipoDocumento,
                    doc.CodigoCliente,
                    doc.Neto,
                    Impuesto1 = doc.Impuesto1,
                    Impuesto2 = doc.Impuesto2,
                    Impuesto3 = doc.Impuesto3,
                    doc.Venta,
                    doc.Descuento,
                    doc.Propina,
                    doc.Total,
                    doc.Abono,
                    doc.Vuelto,
                    Estado = doc.Estado,
                    doc.CodigoTurno,
                    doc.CodigoCaja,
                    doc.FechaRegistro,
                    doc.FechaDiaContable
                },
                transaction: transaction,
                cancellationToken: ct));

            const string insertDetalle = """
                INSERT INTO DDOCUMENTO
                (
                    tDocumento, tItem, tCodigoPedido, tCodigoProducto, nCantidad,
                    nPrecioNeto, nPrecioVenta, nImpuesto1, nImpuesto2, nImpuesto3,
                    nVenta, nDescuento
                )
                VALUES
                (
                    @CodigoDocumento, @Item, @CodigoPedido, @CodigoProducto, @Cantidad,
                    @PrecioNeto, @PrecioVenta, @Impuesto1, @Impuesto2, @Impuesto3,
                    @Venta, @Descuento
                )
                """;

            foreach (var detalle in detalles)
            {
                await connection.ExecuteAsync(new CommandDefinition(
                    insertDetalle,
                    new
                    {
                        detalle.CodigoDocumento,
                        detalle.Item,
                        detalle.CodigoPedido,
                        detalle.CodigoProducto,
                        detalle.Cantidad,
                        detalle.PrecioNeto,
                        detalle.PrecioVenta,
                        detalle.Impuesto1,
                        detalle.Impuesto2,
                        detalle.Impuesto3,
                        detalle.Venta,
                        detalle.Descuento
                    },
                    transaction: transaction,
                    cancellationToken: ct));
            }

            transaction.Commit();
            return cabecera > 0;
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
    }

    public async Task<bool> MarcarCobradoAsync(string codigo, decimal abono, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MDOCUMENTO
            SET tEstadoDocumento = 'CO',
                nAbono = @Abono,
                nVuelto = CASE WHEN @Abono > ISNULL(nTotal, 0) THEN @Abono - ISNULL(nTotal, 0) ELSE 0 END,
                fPago = GETDATE()
            WHERE tDocumento = @Codigo
            """;

        var rows = await connection.ExecuteAsync(new CommandDefinition(
            sql,
            new { Codigo = codigo, Abono = abono },
            cancellationToken: ct));

        return rows > 0;
    }

    public async Task<bool> AnularAsync(string codigo, string usuario, string? motivo = null, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MDOCUMENTO
            SET tEstadoDocumento = 'AN',
                tUsuarioAnulado = @Usuario,
                fRegistroAnulado = GETDATE(),
                tObservacion = CASE
                    WHEN @Motivo IS NULL OR LTRIM(RTRIM(@Motivo)) = '' THEN tObservacion
                    ELSE @Motivo
                END
            WHERE tDocumento = @Codigo
            """;

        var rows = await connection.ExecuteAsync(new CommandDefinition(
            sql,
            new { Codigo = codigo, Usuario = usuario, Motivo = motivo },
            cancellationToken: ct));

        return rows > 0;
    }

    /// <inheritdoc/>
    /// Legacy: SELECT DPEDIDO+TPRODUCTO WHERE tCodigoPedido=@p AND (ISNULL(tFacturado,'0')='0' OR LEN(LTRIM(tFacturado))=0) AND tEstadoItem='N'. BR-DOC-001.
    public async Task<IReadOnlyList<ItemPendienteFacturacionDto>> ObtenerItemsPendientesFacturacionAsync(
        string codigoPedido, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                dp.tCodigoPedido AS CodigoPedido,
                dp.tItem AS Item,
                dp.tCodigoProducto AS CodigoProducto,
                ISNULL(tp.tResumido, dp.tCodigoProducto) AS NombreProducto,
                dp.nPrecioVenta AS PrecioVenta,
                dp.nCantidad AS Cantidad,
                dp.nVenta AS Venta,
                dp.tArea AS Area
            FROM DPEDIDO dp
            INNER JOIN TPRODUCTO tp ON tp.tCodigoProducto = dp.tCodigoProducto
            WHERE dp.tCodigoPedido = @Pedido
              AND (ISNULL(dp.tFacturado, '0') = '0' OR LEN(LTRIM(ISNULL(dp.tFacturado,''))) = 0)
              AND dp.tEstadoItem = 'N'
            ORDER BY dp.nOrden, dp.tItem
            """;

        var rows = await connection.QueryAsync<ItemPendienteFacturacionDto>(
            new CommandDefinition(sql, new { Pedido = codigoPedido }, cancellationToken: ct));

        return rows.ToList().AsReadOnly();
    }

    /// <inheritdoc/>
    /// Legacy: SELECT MDOCUMENTO+TCLIENTE+TMESA+MPEDIDO WHERE tEstadoDocumento='01' AND tCaja=@caja. BR-DOC-008.
    public async Task<IReadOnlyList<DocumentoPendienteDto>> ObtenerDocumentosPendientesCajaAsync(
        string codigoCaja, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                md.tDocumento AS CodigoDocumento,
                MAX(tm.tResumido) AS Mesa,
                MAX(vm.Descripcion) AS Motorizado,
                MAX(mp.tObservacion) AS Observacion,
                MAX(dd.tCodigoPedido) AS CodigoPedido,
                md.nVenta AS Total,
                MAX(md.fRegistro) AS FechaRegistro,
                MAX(tc.tEmpresa) AS Cliente,
                MAX(mp.tTipoPedido) AS TipoPedido,
                md.tTipoDocumento AS TipoDocumento,
                md.tEstadoDocumento AS EstadoDocumento
            FROM MDOCUMENTO md
            LEFT JOIN TCLIENTE tc ON md.tCodigoCliente = tc.tCodigoCliente
            LEFT JOIN DDOCUMENTO dd ON dd.tDocumento = md.tDocumento
            LEFT JOIN MPEDIDO mp ON mp.tCodigoPedido = dd.tCodigoPedido
            LEFT JOIN TMESA tm ON tm.tCodigoMesa = mp.tMesa
            LEFT JOIN vMotorizado vm ON vm.Codigo = mp.tMotorizado
            WHERE md.tEstadoDocumento = '01'
              AND md.tCaja = @Caja
            GROUP BY md.tDocumento, md.nVenta, md.tTipoDocumento, md.tEstadoDocumento
            ORDER BY md.tDocumento
            """;

        var rows = await connection.QueryAsync<DocumentoPendienteDto>(
            new CommandDefinition(sql, new { Caja = codigoCaja }, cancellationToken: ct));

        return rows.ToList().AsReadOnly();
    }

    /// <inheritdoc/>
    /// Legacy: frmDocumento CmdOpcion 7 — EXEC usp_Inforest_Impresion @doc, @modo. BR-DOC-005.
    public async Task<bool> ReimprimirAsync(string codigoDocumento, int modo = 3, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        await connection.ExecuteAsync(new CommandDefinition(
            "usp_Inforest_Impresion",
            new { tDocumento = codigoDocumento, nModo = modo },
            commandType: System.Data.CommandType.StoredProcedure,
            cancellationToken: ct));
        return true;
    }

    private sealed record DocumentoRow(
        string CodigoDocumento,
        string TipoDocumento,
        string? CodigoCliente,
        decimal Neto,
        decimal Impuesto1,
        decimal Impuesto2,
        decimal Impuesto3,
        decimal Venta,
        decimal Descuento,
        decimal Propina,
        decimal Total,
        decimal Abono,
        decimal Vuelto,
        string Estado,
        string? CodigoTurno,
        string? CodigoCaja,
        DateTime? FechaPago,
        DateTime FechaRegistro,
        DateTime? FechaDiaContable,
        string? UsuarioAnulado,
        string? MotivoAnulacion,
        DateTime? FechaAnulacion);
}
