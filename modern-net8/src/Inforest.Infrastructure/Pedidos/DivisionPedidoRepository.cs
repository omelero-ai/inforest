using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Infrastructure.Pedidos;

/// <summary>
/// Repositorio de persistencia para la División de Pedidos.
/// Legacy: operaciones SQL directas de frmDivision.frm sobre DPEDIDO, CPEDIDO, MPEDIDO.
/// </summary>
internal sealed class DivisionPedidoRepository : IDivisionPedidoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public DivisionPedidoRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<IEnumerable<ItemDivisionRecord>> CargarItemsPedidoAsync(
        string codigoPedido,
        CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            SELECT
                d.tCodigoPedido        AS CodigoPedido,
                d.tItem                AS Item,
                d.tTipoPedido          AS TipoPedido,
                d.tCodigoProducto      AS CodigoProducto,
                ISNULL(p.tResumido,'') AS NombreProducto,
                d.tCodigoGrupo         AS CodigoGrupo,
                d.tCodigoSubGrupo      AS CodigoSubGrupo,
                d.tMoneda              AS Moneda,
                ISNULL(d.nPrecioNeto,0)       AS PrecioNeto,
                ISNULL(d.nPrecioImpuesto1,0)  AS PrecioImpuesto1,
                ISNULL(d.nPrecioImpuesto2,0)  AS PrecioImpuesto2,
                ISNULL(d.nPrecioImpuesto3,0)  AS PrecioImpuesto3,
                ISNULL(d.nPrecioVenta,0)      AS PrecioVenta,
                ISNULL(d.nRecargo,0)          AS Recargo,
                ISNULL(d.nDescuento,0)        AS Descuento,
                ISNULL(d.nPrecioOficial,0)    AS PrecioOficial,
                ISNULL(d.nCantidad,0)         AS Cantidad,
                ISNULL(d.nImpuesto1,0)        AS Impuesto1,
                ISNULL(d.nImpuesto2,0)        AS Impuesto2,
                ISNULL(d.nImpuesto3,0)        AS Impuesto3,
                ISNULL(d.nVenta,0)            AS Venta,
                d.tObservacion   AS Observacion,
                d.tCortesia      AS Cortesia,
                ISNULL(d.lImprime,0) AS Imprime,
                d.tEstadoItem    AS EstadoItem,
                d.tArea          AS Area
            FROM DPEDIDO d
            LEFT JOIN TPRODUCTO p ON p.tCodigoProducto = d.tCodigoProducto
            WHERE d.tCodigoPedido = @codigoPedido
            ORDER BY d.tItem
            """;

        var records = await connection.QueryAsync<ItemDivisionRecord>(sql, new { codigoPedido });
        return records;
    }

    public async Task<string?> ObtenerEstadoPedidoAsync(string codigoPedido, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        return await connection.ExecuteScalarAsync<string>(
            "SELECT tEstadoPedido FROM MPEDIDO WHERE tCodigoPedido = @codigoPedido",
            new { codigoPedido });
    }

    public async Task<decimal> ObtenerMontoMaximoPedidoAsync(string codigoPedido, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var result = await connection.ExecuteScalarAsync<decimal?>(
            "SELECT ISNULL(nMontoMaximo, 0) FROM MPEDIDO WHERE tCodigoPedido = @codigoPedido",
            new { codigoPedido });
        return result ?? 0m;
    }

    public async Task<decimal> ObtenerVentaActualPedidoAsync(string codigoPedido, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var result = await connection.ExecuteScalarAsync<decimal?>(
            "SELECT ISNULL(SUM(nVenta), 0) FROM DPEDIDO WHERE tCodigoPedido = @codigoPedido",
            new { codigoPedido });
        return result ?? 0m;
    }

    public async Task<string> ObtenerTipoPedidoAsync(string codigoPedido, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var tipo = await connection.ExecuteScalarAsync<string>(
            "SELECT ISNULL(tTipoPedido,'') FROM MPEDIDO WHERE tCodigoPedido = @codigoPedido",
            new { codigoPedido });
        return tipo ?? string.Empty;
    }

    public async Task<string> GenerarSiguienteCorrelativoPedidoAsync(CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        // Legacy: SELECT max(tCodigoPedido) WHERE substring(tCodigoPedido,1,2) = YY
        const string sql = """
            DECLARE @sAnio CHAR(2) = RIGHT(CONVERT(CHAR(4), YEAR(GETDATE())), 2);
            DECLARE @sCorrela CHAR(10);
            SELECT @sCorrela = MAX(tCodigoPedido)
            FROM MPEDIDO
            WHERE SUBSTRING(tCodigoPedido, 1, 2) = @sAnio;

            IF @sCorrela IS NULL OR LEFT(@sCorrela, 2) <> @sAnio
                SELECT @sAnio + '00000001' AS Codigo
            ELSE
                SELECT @sAnio + RIGHT('00000000' + CAST(CAST(SUBSTRING(@sCorrela, 3, 8) AS INT) + 1 AS VARCHAR), 8) AS Codigo
            """;
        var codigo = await connection.ExecuteScalarAsync<string>(sql);
        return codigo ?? throw new InvalidOperationException("No se pudo generar correlativo para MPEDIDO.");
    }

    public async Task<bool> TieneItemPropiedadAsync(string codigoPedido, string item, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var count = await connection.ExecuteScalarAsync<int>(
            "SELECT COUNT(*) FROM TPRODUCTOPROPIEDAD WHERE tCodigoPedido = @codigoPedido AND tItem = @item",
            new { codigoPedido, item });
        return count > 0;
    }

    public async Task ConfirmarDivisionAsync(ConfirmarDivisionParams p, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        using var transaction = connection.BeginTransaction();

        try
        {
            // 1. DELETE + INSERT DPEDIDO para el pedido origen (items que permanecen)
            await connection.ExecuteAsync(
                "DELETE DPEDIDO WHERE tCodigoPedido = @origen",
                new { origen = p.CodigoPedidoOrigen }, transaction);

            await InsertarItemsDpedidoAsync(connection, transaction, p.CodigoPedidoOrigen, p.ItemsOrigen, 0);

            // 2. spUpdate_DPEDIDO_Ina para el origen
            await _spExecutor.ExecuteAsync(connection, "spUpdate_DPEDIDO_Ina",
                new { pedido = p.CodigoPedidoOrigen }, transaction, cancellationToken);

            // 3. DELETE + INSERT CPEDIDO para el origen
            await connection.ExecuteAsync(
                "DELETE CPEDIDO WHERE tCodigoPedido = @origen",
                new { origen = p.CodigoPedidoOrigen }, transaction);

            // 4. Para cada pedido destino con items
            foreach (var (destino, items) in p.PedidosDestino)
            {
                var codigoFinal = destino.CodigoDefinitivo ?? destino.CodigoTemporal;

                if (destino.CodigoDefinitivo is not null)
                {
                    // Crear MPEDIDO nuevo copiando del origen
                    await CrearMpedidoFromOrigenAsync(connection, transaction,
                        p.CodigoPedidoOrigen, codigoFinal, destino.Observacion, p.FechaDiaContable);
                }

                // INSERT DPEDIDO para el pedido destino
                int offsetItem = await ObtenerMaxItemAsync(connection, transaction, codigoFinal);
                await InsertarItemsDpedidoAsync(connection, transaction, codigoFinal, items, offsetItem);

                // spUpdate_DPEDIDO_Ina para el destino
                await _spExecutor.ExecuteAsync(connection, "spUpdate_DPEDIDO_Ina",
                    new { pedido = codigoFinal }, transaction, cancellationToken);
            }

            transaction.Commit();
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
    }

    // ─── Helpers privados ──────────────────────────────────────────────────────

    private static async Task InsertarItemsDpedidoAsync(
        IDbConnection connection,
        IDbTransaction transaction,
        string codigoPedido,
        IEnumerable<ItemDivision> items,
        int itemOffset)
    {
        const string sql = """
            INSERT INTO DPEDIDO (
                tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo,
                tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta,
                nRecargo, nDescuento, nPrecioOficial, nCantidad,
                nImpuesto1, nImpuesto2, nImpuesto3, nVenta,
                tObservacion, tCortesia, lImprime, tEstadoItem, tArea
            ) VALUES (
                @CodigoPedido, @Item, @TipoPedido, @CodigoProducto, @CodigoGrupo, @CodigoSubGrupo,
                @Moneda, @PrecioNeto, @PrecioImpuesto1, @PrecioImpuesto2, @PrecioImpuesto3, @PrecioVenta,
                @Recargo, @Descuento, @PrecioOficial, @Cantidad,
                @Impuesto1, @Impuesto2, @Impuesto3, @Venta,
                @Observacion, @Cortesia, @Imprime, @EstadoItem, @Area
            )
            """;

        int idx = 0;
        foreach (var item in items)
        {
            idx++;
            string tItem = (int.Parse(item.Item) + itemOffset).ToString().PadLeft(3, '0');
            await connection.ExecuteAsync(sql, new
            {
                CodigoPedido = codigoPedido,
                Item = tItem,
                item.TipoPedido,
                item.CodigoProducto,
                item.CodigoGrupo,
                item.CodigoSubGrupo,
                item.Moneda,
                item.PrecioNeto,
                item.PrecioImpuesto1,
                item.PrecioImpuesto2,
                item.PrecioImpuesto3,
                item.PrecioVenta,
                item.Recargo,
                item.Descuento,
                item.PrecioOficial,
                item.Cantidad,
                item.Impuesto1,
                item.Impuesto2,
                item.Impuesto3,
                item.Venta,
                item.Observacion,
                item.Cortesia,
                item.Imprime,
                item.EstadoItem,
                item.Area
            }, transaction);
        }
    }

    private static async Task CrearMpedidoFromOrigenAsync(
        IDbConnection connection,
        IDbTransaction transaction,
        string codigoPedidoOrigen,
        string codigoNuevo,
        string observacion,
        string fechaDiaContable)
    {
        // Full MPEDIDO INSERT — same SELECT...FROM MPEDIDO pattern as Legacy VB6 frmDivision.frm.
        // Copies all columns from the origin pedido, overriding key fields for the new pedido.
        const string sqlFull = """
            INSERT INTO MPEDIDO (
                tCodigoPedido, nCorrelativo, tClienteDelivery, tClienteCtaCte,
                fFecha, tMoneda, nMonto, tEstadoPedido, tTipoAtencion, tTipoPedido,
                lPrioridad, tAnulacionPedido, tMesa, nMesa, tMozo, tMotorizado,
                tCaja, tSalon, tTurno, fProgramacion, nTiempo, tObservacion,
                tUsuario, fRegistro, nAdulto, nNino,
                tMotivoAnulacion, tUsuarioAnulado, fRegAnulado, tObservacionAnulado,
                tTurnoAnulado, tClienteCorp, tTienda, fRegCuenta, nPrecuenta,
                tCajaAnterior, tTurnoAnterior, tComanda, tPuntoVenta,
                tHabitacion, tReserva, tPasajero, tCompania, tContacto, tFichaPasajero,
                tTipoComanda, nDescuento, tDescuento, tObservacionDescuento, tUsuarioDescuento,
                tEmpacador, fEmpacador, fAsignacion, fSalida, fLlegada,
                fdiacontable, tCodigoInvitado, tCodigoPariente,
                lAtendidoC, fAtendidoC, tUsuarioAtendio, tEntregara, tTipoRecepcion, nMinutosAntesEnvio
            )
            SELECT
                @codigoNuevo, nCorrelativo, tClienteDelivery, tClienteCtaCte,
                GETDATE(), tMoneda, nMonto, tEstadoPedido, tTipoAtencion, tTipoPedido,
                lPrioridad, tAnulacionPedido, '' AS tMesa, 0 AS nMesa, tMozo, tMotorizado,
                tCaja, tSalon, tTurno, fProgramacion, nTiempo, @observacion,
                tUsuario, GETDATE(), 0, 0,
                tMotivoAnulacion, tUsuarioAnulado, fRegAnulado, tObservacionAnulado,
                tTurnoAnulado, tClienteCorp, tTienda, fRegCuenta, 0 AS nPrecuenta,
                tCajaAnterior, tTurnoAnterior, '' AS tComanda, tPuntoVenta,
                tHabitacion, tReserva, tPasajero, tCompania, tContacto, tFichaPasajero,
                tTipoComanda, nDescuento, tDescuento, tObservacionDescuento, tUsuarioDescuento,
                tEmpacador, fEmpacador, fAsignacion, fSalida, fLlegada,
                @fechaDiaContable,
                ISNULL(tCodigoInvitado,''), ISNULL(tCodigoPariente,''),
                lAtendidoC, fAtendidoC, tUsuarioAtendio, tEntregara, 0, nMinutosAntesEnvio
            FROM MPEDIDO
            WHERE tCodigoPedido = @codigoPedidoOrigen
            """;

        await connection.ExecuteAsync(sqlFull, new
        {
            codigoNuevo,
            observacion,
            fechaDiaContable,
            codigoPedidoOrigen
        }, transaction);
    }

    private static async Task<int> ObtenerMaxItemAsync(
        IDbConnection connection,
        IDbTransaction transaction,
        string codigoPedido)
    {
        var max = await connection.ExecuteScalarAsync<int?>(
            "SELECT ISNULL(MAX(CAST(tItem AS INT)), 0) FROM DPEDIDO WHERE tCodigoPedido = @codigoPedido",
            new { codigoPedido }, transaction);
        return max ?? 0;
    }
}
