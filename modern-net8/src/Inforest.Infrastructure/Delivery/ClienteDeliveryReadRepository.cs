using Dapper;
using Inforest.Application.Delivery;
using Inforest.Application.Interfaces;

namespace Inforest.Infrastructure.Delivery;

/// <summary>
/// Repositorio de lectura Dapper para búsqueda y estadísticas de clientes delivery.
/// <para>
/// Legacy: <c>frmBusquedaDelivery.frm</c> — POS-FUNC-036.
/// Tablas/Vistas: <c>TDELIVERY</c>, <c>vZona</c>, <c>vDelivery</c>,
/// <c>MPEDIDO</c>, <c>DPEDIDO</c>, <c>MDOCUMENTO</c>, <c>vCliente</c>, <c>vTienda</c>.
/// </para>
/// </summary>
internal sealed class ClienteDeliveryReadRepository : IClienteDeliveryReadRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public ClienteDeliveryReadRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    /// <inheritdoc />
    public async Task<IReadOnlyList<ClienteDeliveryListadoItem>> ListarMantenimientoAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT Codigo                        AS Codigo,
                   TipoCliente                   AS TipoCliente,
                   Cliente                       AS Cliente,
                   tTelefono                     AS Telefono,
                   tDireccion                    AS Direccion,
                   ISNULL(nLinea, 0)             AS Linea,
                   ISNULL(nDescuento, 0)         AS Descuento,
                   EstadoFrecuente               AS EstadoFrecuente,
                   CAST(ISNULL(lActivo, 0) AS bit) AS Activo
            FROM dbo.vDelivery
            ORDER BY Cliente, Codigo
            """;

        var rows = await conn.QueryAsync<ClienteDeliveryListadoItem>(sql);
        return rows.AsList();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<ClienteDeliveryBusquedaItem>> ListarActivosConZonaAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT d.tCodigoDelivery AS CodigoDelivery,
                   d.tTelefono       AS Telefono,
                   LTRIM(ISNULL(d.tApellido,'')) + ' ' + LTRIM(ISNULL(d.tNombre,'')) AS Cliente,
                   ISNULL(d.nDescuento, 0)   AS Descuento,
                   LTRIM(ISNULL(d.tReferencia,''))              AS Referencia,
                   z.Descripcion             AS Zona,
                   ISNULL(d.nAcumulado, 0)   AS Acumulado,
                   ISNULL(d.nUtilizado, 0)   AS Utilizado,
                   ISNULL(d.nDisponible, 0)  AS Disponible
            FROM   dbo.TDELIVERY d
                   LEFT JOIN dbo.vZona z ON d.tZona = z.Codigo
            WHERE  d.lActivo = 1
            ORDER BY Cliente
            """;
        var rows = await conn.QueryAsync<ClienteDeliveryBusquedaItem>(sql);
        return rows.AsList();
    }

    /// <inheritdoc />
    public async Task<ClienteDeliveryDetalleBusqueda?> ObtenerDetalleAsync(string codigoDelivery, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT Codigo           AS CodigoDelivery,
                   tApellido        AS Apellido,
                   tNombre          AS Nombre,
                   tTelefono        AS Telefono,
                   tDireccion       AS Direccion,
                   Zona             AS Zona,
                   tReferencia      AS Referencia,
                   tObservacion     AS Observacion,
                   ISNULL(nDescuento, 0)  AS Descuento,
                   ISNULL(nAcumulado, 0)  AS Acumulado,
                   ISNULL(nUtilizado, 0)  AS Utilizado,
                   ISNULL(nDisponible, 0) AS Disponible,
                   TipoCliente      AS TipoCliente
            FROM   dbo.vDelivery
            WHERE  Codigo = @CodigoDelivery
            """;
        return await conn.QueryFirstOrDefaultAsync<ClienteDeliveryDetalleBusqueda>(
            sql, new { CodigoDelivery = codigoDelivery });
    }

    /// <inheritdoc />
    public async Task<EstadisticasClienteDelivery> ObtenerEstadisticasAsync(
        string codigoDelivery, int diasHistorico, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);

        // ── Pedidos activos en los últimos N días ────────────────────────────
        // Legacy: frmBusquedaDelivery.frm cmdOpcion(3) bloque "Segundo Calculo"
        const string sqlPedidos = """
            SELECT COUNT(DISTINCT mp.tCodigoPedido) AS NumeroPedidos,
                   ISNULL(SUM(dp.nVenta), 0)        AS TotalVenta
            FROM   dbo.MPEDIDO mp
                   INNER JOIN dbo.DPEDIDO dp ON mp.tCodigoPedido = dp.tCodigoPedido
            WHERE  mp.tEstadoPedido <> '03'
              AND  dp.tEstadoItem   = 'N'
              AND  mp.tTipoPedido   = '02'
              AND  mp.tClienteDelivery = @CodigoDelivery
              AND  mp.fRegistro >= DATEADD(day, -@Dias, GETDATE())
            """;

        var pedidoStats = await conn.QueryFirstOrDefaultAsync<(int NumeroPedidos, decimal TotalVenta)>(
            sqlPedidos, new { CodigoDelivery = codigoDelivery, Dias = diasHistorico });

        var desde = diasHistorico > 0
            ? DateOnly.FromDateTime(DateTime.Today.AddDays(-diasHistorico))
            : (DateOnly?)null;

        // ── Último pedido del cliente ─────────────────────────────────────────
        // Legacy: frmBusquedaDelivery.frm — MAX(tCodigoPedido) + join doc
        const string sqlUltimoPedido = """
            SELECT TOP 1 tCodigoPedido AS Codigo
            FROM   dbo.MPEDIDO
            WHERE  tClienteDelivery = @CodigoDelivery
              AND  tEstadoPedido   <> '03'
            ORDER BY tCodigoPedido DESC
            """;

        var ultimoPedidoCodigo = await conn.QueryFirstOrDefaultAsync<string?>(
            sqlUltimoPedido, new { CodigoDelivery = codigoDelivery });

        if (string.IsNullOrEmpty(ultimoPedidoCodigo))
        {
            return new EstadisticasClienteDelivery(
                desde, pedidoStats.NumeroPedidos, pedidoStats.TotalVenta,
                null, null, 0m, null, null);
        }

        // Legacy: join DPEDIDO - MDOCUMENTO - vCliente para el último pedido
        const string sqlUltimoDoc = """
            SELECT SUM(dp.nVenta)              AS MontoUltimoDoc,
                   MAX(md.fRegistro)           AS FechaUltimaCompra,
                   MAX(md.tDocumento)          AS UltimoDocumento,
                   MAX(vc.Descripcion)         AS RazonSocial,
                   MAX(vc.tIdentidad)          AS Ruc
            FROM   dbo.DPEDIDO dp
                   LEFT JOIN dbo.MDOCUMENTO md ON md.tDocumento = dp.tDocumento
                   LEFT JOIN dbo.vCliente   vc ON vc.Codigo = md.tCodigoCliente
            WHERE  dp.tCodigoPedido = @CodigoPedido
            GROUP BY dp.tCodigoPedido
            """;

        var docStats = await conn.QueryFirstOrDefaultAsync<(decimal MontoUltimoDoc, DateTime? FechaUltimaCompra, string? UltimoDocumento, string? RazonSocial, string? Ruc)>(
            sqlUltimoDoc, new { CodigoPedido = ultimoPedidoCodigo });

        return new EstadisticasClienteDelivery(
            desde,
            pedidoStats.NumeroPedidos,
            pedidoStats.TotalVenta,
            docStats.FechaUltimaCompra,
            docStats.UltimoDocumento,
            docStats.MontoUltimoDoc,
            docStats.RazonSocial,
            docStats.Ruc);
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<TiendaDeliveryItem>> ObtenerTiendasAsync(string codigoDelivery, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT Codigo       AS CodigoTienda,
                   Descripcion  AS Descripcion,
                   tDireccion   AS Direccion
            FROM   dbo.vTienda
            WHERE  lActivo = 1
              AND  tCodigoDelivery = @CodigoDelivery
            ORDER BY Descripcion
            """;
        var rows = await conn.QueryAsync<TiendaDeliveryItem>(sql, new { CodigoDelivery = codigoDelivery });
        return rows.AsList();
    }
}
