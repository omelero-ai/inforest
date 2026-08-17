using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;

namespace Inforest.Infrastructure.Ventas;

/// <summary>
/// Repositorio Dapper para actualizar cortesía y canal de venta del pedido.
/// Legacy: FrmActualizarPedidos — usp_ActualizarCabPeDoc. BR-ACTPED-001/002/003.
/// Tablas/SPs: MDOCUMENTO, DDOCUMENTO, MPEDIDO, usp_ActualizarCabPeDoc.
/// </summary>
internal sealed class ActualizarDatosPedidoRepository : IActualizarDatosPedidoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public ActualizarDatosPedidoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    /// <inheritdoc />
    public async Task<DatosPedidoActualizar?> ObtenerDatosActualesAsync(
        string codigoDocumento, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT DISTINCT
                t2.tCodigoPedido,
                ISNULL(t1.tCortesia, '')           AS CodigoCortesia,
                ISNULL(t5.Descripcion, '')          AS DescripcionCortesia,
                ISNULL(t3.tTipoPedido, '')          AS CodigoCanalVenta,
                ISNULL(t4.Descripcion, '')          AS DescripcionCanalVenta
            FROM MDOCUMENTO t1
            LEFT JOIN DDOCUMENTO t2     ON t1.tDocumento = t2.tDocumento
            LEFT JOIN MPEDIDO t3        ON t2.tCodigoPedido = t3.tCodigoPedido
            LEFT JOIN vTipoPedido t4    ON t3.tTipoPedido = t4.Codigo
            LEFT JOIN vCortesia t5      ON t5.Codigo = t1.tCortesia
            WHERE t1.tDocumento = @Doc
            """;

        var row = await connection.QueryFirstOrDefaultAsync<DatosRow>(
            new CommandDefinition(sql, new { Doc = codigoDocumento }, cancellationToken: ct));

        if (row is null) return null;
        return new DatosPedidoActualizar(
            row.tCodigoPedido ?? string.Empty,
            row.CodigoCortesia,
            row.DescripcionCortesia,
            row.CodigoCanalVenta,
            row.DescripcionCanalVenta);
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<OpcionCatalogo>> ObtenerCortesiasAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = "SELECT Codigo, Descripcion FROM vCortesia WHERE lActivo = 1 ORDER BY Codigo";
        var rows = await connection.QueryAsync<CatalogoRow>(
            new CommandDefinition(sql, cancellationToken: ct));
        return rows.Select(r => new OpcionCatalogo(r.Codigo, r.Descripcion)).ToList();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<OpcionCatalogo>> ObtenerCanalesVentaAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = "SELECT Codigo, Descripcion FROM vTipoPedido WHERE lActivo = 1 ORDER BY Codigo";
        var rows = await connection.QueryAsync<CatalogoRow>(
            new CommandDefinition(sql, cancellationToken: ct));
        return rows.Select(r => new OpcionCatalogo(r.Codigo, r.Descripcion)).ToList();
    }

    /// <inheritdoc />
    public async Task ActualizarCortesiaAsync(
        string codigoDocumento,
        string codigoCortesiaNueva,
        string codigoCortesiaAnterior,
        string codigoUsuario,
        CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        await connection.ExecuteAsync(
            new CommandDefinition(
                "usp_ActualizarCabPeDoc",
                new
                {
                    opcion = "1",
                    param1 = codigoDocumento,
                    param2 = codigoCortesiaNueva,
                    param3 = codigoCortesiaAnterior,
                    param4 = codigoUsuario,
                    param5 = string.Empty,
                    param6 = string.Empty
                },
                commandType: System.Data.CommandType.StoredProcedure,
                cancellationToken: ct));
    }

    /// <inheritdoc />
    public async Task ActualizarCanalVentaAsync(
        string codigoDocumento,
        string codigoCanalNuevo,
        string codigoClienteDelivery,
        string codigoCanalAnterior,
        string codigoUsuario,
        CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        await connection.ExecuteAsync(
            new CommandDefinition(
                "usp_ActualizarCabPeDoc",
                new
                {
                    opcion = "2",
                    param1 = codigoDocumento,
                    param2 = codigoCanalNuevo,
                    param3 = codigoClienteDelivery,
                    param4 = codigoCanalAnterior,
                    param5 = codigoUsuario,
                    param6 = string.Empty
                },
                commandType: System.Data.CommandType.StoredProcedure,
                cancellationToken: ct));
    }

    // ── Private row types ────────────────────────────────────────────────────
#pragma warning disable CS8618
    private sealed class DatosRow
    {
        public string? tCodigoPedido { get; set; }
        public string CodigoCortesia { get; set; }
        public string DescripcionCortesia { get; set; }
        public string CodigoCanalVenta { get; set; }
        public string DescripcionCanalVenta { get; set; }
    }

    private sealed class CatalogoRow
    {
        public string Codigo { get; set; }
        public string Descripcion { get; set; }
    }
#pragma warning restore CS8618
}
