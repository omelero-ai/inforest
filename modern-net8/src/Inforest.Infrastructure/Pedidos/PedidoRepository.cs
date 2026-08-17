using System.Reflection;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Pedidos;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;

namespace Inforest.Infrastructure.Pedidos;

internal sealed class PedidoRepository : IPedidoRepository, IPedidoReadRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public PedidoRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<Pedido?> ObtenerPorCodigoAsync(string codigoPedido, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            SELECT
                tCodigoPedido AS CodigoPedido,
                NULLIF(tMesa, '') AS CodigoMesa,
                NULLIF(tTurno, '') AS CodigoTurno,
                NULLIF(tUsuario, '') AS CodigoUsuario,
                NULLIF(tTipoPedido, '') AS TipoPedido,
                NULLIF(tEstadoPedido, '') AS EstadoPedido,
                fFecha AS FechaRegistro,
                fDiaContable AS FechaDiaContable,
                ISNULL(nAdulto, 0) AS NumeroAdultos,
                ISNULL(nNino, 0) AS NumeroNinos,
                tObservacion AS Observacion
            FROM MPEDIDO
            WHERE tCodigoPedido = @codigo
            """;

        var record = await connection.QueryFirstOrDefaultAsync<PedidoRecord>(sql, new { codigo = codigoPedido });
        return record is null ? null : Map(record);
    }

    public async Task<IEnumerable<Pedido>> ObtenerAbiertosAsync(string codigoCaja, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            SELECT
                tCodigoPedido AS CodigoPedido,
                NULLIF(tMesa, '') AS CodigoMesa,
                NULLIF(tTurno, '') AS CodigoTurno,
                NULLIF(tUsuario, '') AS CodigoUsuario,
                NULLIF(tTipoPedido, '') AS TipoPedido,
                NULLIF(tEstadoPedido, '') AS EstadoPedido,
                fFecha AS FechaRegistro,
                fDiaContable AS FechaDiaContable,
                ISNULL(nAdulto, 0) AS NumeroAdultos,
                ISNULL(nNino, 0) AS NumeroNinos,
                tObservacion AS Observacion
            FROM MPEDIDO
            WHERE tCaja = @caja AND ISNULL(tEstadoPedido, '01') NOT IN ('AN', '03', 'CE')
            ORDER BY fFecha DESC
            """;

        var records = await connection.QueryAsync<PedidoRecord>(sql, new { caja = codigoCaja });
        return records.Select(Map).ToList();
    }

    public async Task InsertarAsync(Pedido pedido, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var metadata = PedidoRepositoryContext.Current;
        var parameters = new DynamicParameters();
        parameters.Add("@tCliente", string.Empty);
        parameters.Add("@tTipoPedido", LegacyMappings.MapCanal(metadata?.Canal ?? pedido.CanalVenta));
        parameters.Add("@lPrioridad", false);
        parameters.Add("@tTipoAtencion", "01");
        parameters.Add("@tMesa", metadata?.CodigoMesa ?? pedido.CodigoMesa ?? string.Empty);
        parameters.Add("@tMozo", metadata?.CodigoMozo ?? string.Empty);
        parameters.Add("@tMotorizado", string.Empty);
        parameters.Add("@tCaja", metadata?.CodigoCaja ?? string.Empty);
        parameters.Add("@tSalon", string.Empty);
        parameters.Add("@tTurno", metadata?.CodigoTurno ?? pedido.CodigoTurno ?? string.Empty);
        parameters.Add("@tObservacion", metadata?.Observacion ?? pedido.Observacion ?? string.Empty);
        parameters.Add("@nTiempo", 0);
        parameters.Add("@tUsuario", metadata?.CodigoUsuario ?? pedido.CodigoUsuario ?? string.Empty);
        parameters.Add("@nAdulto", metadata?.NumeroAdultos ?? pedido.NumeroPersonas);
        parameters.Add("@nNino", metadata?.NumeroNinos ?? 0);
        parameters.Add("@nMesa", 0);
        parameters.Add("@tPuntoVenta", LegacyMappings.MapCanal(metadata?.Canal ?? pedido.CanalVenta));
        parameters.Add("@tHabitacion", string.Empty);
        parameters.Add("@tReserva", string.Empty);
        parameters.Add("@tPasajero", string.Empty);
        parameters.Add("@tCompania", string.Empty);
        parameters.Add("@tContacto", string.Empty);
        parameters.Add("@nDescuento", 0m);
        parameters.Add("@tDescuento", string.Empty);
        parameters.Add("@tObservacionDescuento", string.Empty);
        parameters.Add("@tAutorizaDescuento", string.Empty);
        parameters.Add("@nTiempoDelivery", 0);
        parameters.Add("@tTienda", string.Empty);
        parameters.Add("@fDiaContable", metadata?.FechaDiaContable ?? pedido.FechaDiaContable ?? DateTime.Today);
        parameters.Add("@fProgramacion", null);
        parameters.Add("@tCodigoInvitado", string.Empty);
        parameters.Add("@tcodigoPariente", string.Empty);
        parameters.Add("@tEntregarA", string.Empty);
        parameters.Add("@nTiempoAntesEnvio", 0);
        parameters.Add("@nMontoMaximo", string.Empty);
        parameters.Add("@tPedido", dbType: System.Data.DbType.String, size: 10, direction: System.Data.ParameterDirection.Output);
        parameters.Add("@codigoOrigenVentas", string.Empty);
        parameters.Add("@tTotemMesa", string.Empty);

        await _spExecutor.ExecuteAsync(connection, "spIns_MPEDIDO", parameters, cancellationToken: cancellationToken);

        var codigoGenerado = parameters.Get<string>("@tPedido");
        if (!string.IsNullOrWhiteSpace(codigoGenerado))
            SetProperty(pedido, nameof(Pedido.CodigoPedido), codigoGenerado);
    }

    public async Task ActualizarAsync(Pedido pedido, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var parameters = new DynamicParameters();
        parameters.Add("@tCliente", string.Empty);
        parameters.Add("@tTipoPedido", LegacyMappings.MapCanal(pedido.CanalVenta));
        parameters.Add("@lPrioridad", false);
        parameters.Add("@tTipoAtencion", "01");
        parameters.Add("@tMozo", string.Empty);
        parameters.Add("@tMotorizado", string.Empty);
        parameters.Add("@tObservacion", pedido.Observacion ?? string.Empty);
        parameters.Add("@nTiempo", 0);
        parameters.Add("@tPuntoVenta", LegacyMappings.MapCanal(pedido.CanalVenta));
        parameters.Add("@tHabitacion", string.Empty);
        parameters.Add("@tReserva", string.Empty);
        parameters.Add("@tPasajero", string.Empty);
        parameters.Add("@tCompania", string.Empty);
        parameters.Add("@tContacto", string.Empty);
        parameters.Add("@nDescuento", 0m);
        parameters.Add("@tDescuento", string.Empty);
        parameters.Add("@tObservacionDescuento", string.Empty);
        parameters.Add("@tAutorizaDescuento", string.Empty);
        parameters.Add("@tTienda", string.Empty);
        parameters.Add("@fProgramacion", null);
        parameters.Add("@tcodigoInvitado", string.Empty);
        parameters.Add("@tCodigoPariente", string.Empty);
        parameters.Add("@tEntregarA", string.Empty);
        parameters.Add("@nTiempoAntesEnvio", 0);
        parameters.Add("@nMontoMaximo", string.Empty);
        parameters.Add("@tPedido", pedido.CodigoPedido);
        parameters.Add("@codigoOrigenVentas", string.Empty);
        parameters.Add("@tMesa", pedido.CodigoMesa ?? string.Empty);
        parameters.Add("@tTotemMesa", string.Empty);

        await _spExecutor.ExecuteAsync(connection, "spUpd_MPEDIDO", parameters, cancellationToken: cancellationToken);

        const string sqlEstado = """
            UPDATE MPEDIDO
            SET tEstadoPedido = @estado,
                tObservacion = @observacion
            WHERE tCodigoPedido = @codigo
            """;

        await connection.ExecuteAsync(sqlEstado, new
        {
            codigo = pedido.CodigoPedido,
            estado = LegacyMappings.MapEstado(pedido.Estado),
            observacion = pedido.Observacion ?? string.Empty
        });
    }

    public async Task<IReadOnlyList<Pedido>> ObtenerPorTurnoAsync(string codigoTurno, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT
                tCodigoPedido AS CodigoPedido,
                NULLIF(tMesa, '') AS CodigoMesa,
                NULLIF(tTurno, '') AS CodigoTurno,
                NULLIF(tUsuario, '') AS CodigoUsuario,
                NULLIF(tTipoPedido, '') AS TipoPedido,
                NULLIF(tEstadoPedido, '') AS EstadoPedido,
                fFecha AS FechaRegistro,
                fDiaContable AS FechaDiaContable,
                ISNULL(nAdulto, 0) AS NumeroAdultos,
                ISNULL(nNino, 0) AS NumeroNinos,
                tObservacion AS Observacion
            FROM MPEDIDO
            WHERE tTurno = @turno AND ISNULL(tEstadoPedido, '01') <> 'AN'
            ORDER BY fFecha DESC
            """;
        var records = await connection.QueryAsync<PedidoRecord>(sql, new { turno = codigoTurno });
        return records.Select(Map).ToList();
    }

    public async Task<IReadOnlyList<Pedido>> ObtenerPorMesaAsync(string codigoMesa, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT
                tCodigoPedido AS CodigoPedido,
                NULLIF(tMesa, '') AS CodigoMesa,
                NULLIF(tTurno, '') AS CodigoTurno,
                NULLIF(tUsuario, '') AS CodigoUsuario,
                NULLIF(tTipoPedido, '') AS TipoPedido,
                NULLIF(tEstadoPedido, '') AS EstadoPedido,
                fFecha AS FechaRegistro,
                fDiaContable AS FechaDiaContable,
                ISNULL(nAdulto, 0) AS NumeroAdultos,
                ISNULL(nNino, 0) AS NumeroNinos,
                tObservacion AS Observacion
            FROM MPEDIDO
            WHERE tMesa = @mesa AND ISNULL(tEstadoPedido, '01') <> 'AN'
            ORDER BY fFecha DESC
            """;
        var records = await connection.QueryAsync<PedidoRecord>(sql, new { mesa = codigoMesa });
        return records.Select(Map).ToList();
    }

    public async Task<IReadOnlyList<DetallePedido>> ObtenerDetalleAsync(string codigoPedido, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT
                tCodigoPedido AS CodigoPedido,
                tItem AS Item,
                tCodigoProducto AS CodigoProducto,
                NULLIF(tCodigoGrupo, '') AS CodigoGrupo,
                NULLIF(tCodigoSubGrupo, '') AS CodigoSubGrupo,
                CAST(ISNULL(nPrecioNeto, 0) AS decimal(18,2)) AS PrecioNeto,
                CAST(ISNULL(nPrecioVenta, 0) AS decimal(18,2)) AS PrecioVenta,
                CAST(ISNULL(nCantidad, 0) AS decimal(18,2)) AS Cantidad,
                CAST(ISNULL(nImpuesto1, 0) AS decimal(18,2)) AS Impuesto1,
                CAST(ISNULL(nImpuesto2, 0) AS decimal(18,2)) AS Impuesto2,
                CAST(ISNULL(nImpuesto3, 0) AS decimal(18,2)) AS Impuesto3,
                CAST(ISNULL(nRecargo, 0) AS decimal(18,2)) AS Recargo,
                CAST(ISNULL(nDescuento, 0) AS decimal(18,2)) AS Descuento,
                tObservacion AS Observacion,
                NULLIF(tArea, '') AS Area,
                CAST(ISNULL(lImprime, 0) AS bit) AS Imprime,
                NULLIF(tEstadoItem, '') AS EstadoItem
            FROM DPEDIDO
            WHERE tCodigoPedido = @codigo
            ORDER BY tItem
            """;
        var records = await connection.QueryAsync<DetallePedidoRecord>(sql, new { codigo = codigoPedido });
        return records.Select(Map).ToList();
    }

    private static Pedido Map(PedidoRecord record)
    {
        var pedido = Pedido.Crear(
            record.CodigoPedido,
            record.CodigoMesa,
            record.CodigoTurno ?? "LEGACY",
            record.CodigoUsuario ?? string.Empty,
            MapCanal(record.TipoPedido),
            Math.Max(0, record.NumeroAdultos + record.NumeroNinos));

        SetProperty(pedido, nameof(Pedido.Estado), LegacyMappings.MapEstado(record.EstadoPedido));
        SetProperty(pedido, nameof(Pedido.FechaRegistro), record.FechaRegistro);
        SetProperty(pedido, nameof(Pedido.FechaDiaContable), record.FechaDiaContable);
        SetProperty(pedido, nameof(Pedido.Observacion), string.IsNullOrWhiteSpace(record.Observacion) ? null : record.Observacion.Trim());
        return pedido;
    }

    private static DetallePedido Map(DetallePedidoRecord record)
    {
        var detalle = DetallePedido.Crear(
            record.CodigoPedido,
            record.Item,
            record.CodigoProducto,
            record.Cantidad,
            record.PrecioNeto,
            record.PrecioVenta,
            record.Impuesto1,
            record.Impuesto2,
            record.Impuesto3);

        SetProperty(detalle, nameof(DetallePedido.CodigoGrupo), record.CodigoGrupo);
        SetProperty(detalle, nameof(DetallePedido.CodigoSubGrupo), record.CodigoSubGrupo);
        SetProperty(detalle, nameof(DetallePedido.Recargo), record.Recargo);
        SetProperty(detalle, nameof(DetallePedido.Descuento), record.Descuento);
        SetProperty(detalle, nameof(DetallePedido.Observacion), record.Observacion);
        SetProperty(detalle, nameof(DetallePedido.Area), record.Area);
        SetProperty(detalle, nameof(DetallePedido.Imprime), record.Imprime);
        SetProperty(detalle, nameof(DetallePedido.EstadoItem), record.EstadoItem);
        return detalle;
    }

    private static CanalVenta MapCanal(string? tipoPedido)
        => (tipoPedido ?? string.Empty).Trim() switch
        {
            "02" => CanalVenta.Delivery,
            "03" => CanalVenta.Llevar,
            "04" => CanalVenta.Canal4,
            "05" => CanalVenta.Canal5,
            _ => CanalVenta.Local
        };

    private static void SetProperty<T>(object target, string propertyName, T value)
    {
        var property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException($"No se encontró la propiedad {propertyName}.");
        property.SetValue(target, value);
    }

    // ── Vista enriquecida (frmDetallePedido.frm) ─────────────────────────────

    /// <inheritdoc />
    public async Task<IReadOnlyList<ItemPedidoVista>> ObtenerDetalleExtendidoAsync(
        string codigoPedido, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // Legacy: vPedidoDetalle view — DPEDIDO JOIN TPRODUCTO JOIN vCortesia.
        // BR-PEDIDO-005: solo ítems activos (tEstadoItem = 'N').
        const string sql = """
            SELECT
                d.tItem                     AS Item,
                d.tCodigoProducto           AS CodigoProducto,
                ISNULL(p.tDetallado, d.tCodigoProducto) AS Producto,
                CAST(ISNULL(d.nPrecioOficial, 0)  AS decimal(18,2)) AS PrecioOficial,
                CAST(ISNULL(d.nDescuento,    0)   AS decimal(18,2)) AS Descuento,
                CAST(ISNULL(d.nPrecioVenta,  0)   AS decimal(18,2)) AS PrecioVenta,
                CAST(ISNULL(d.nCantidad,     0)   AS decimal(18,4)) AS Cantidad,
                CAST(ISNULL(d.nVenta,        0)   AS decimal(18,2)) AS SubTotal,
                ISNULL(d.tEstadoItem, 'N')         AS EstadoItem,
                ISNULL(d.tFacturado, '')            AS Facturado,
                CAST(ISNULL(d.lImprime, 0)         AS bit)          AS Imprime,
                CAST(0 AS bit)                     AS TienePropiedad,
                CAST(CASE WHEN DATALENGTH(d.tObservacion) > 0 THEN 1 ELSE 0 END AS bit) AS TieneObservacion,
                CAST(ISNULL(d.lCorte, 0)           AS bit)          AS Corte,
                d.tObservacion                     AS Observacion,
                d.tDocumento                       AS Documento,
                d.tusuariod                        AS Usuario,
                d.fEnvio                           AS FechaEnvio,
                d.tComanda                         AS Comanda
            FROM DPEDIDO d
            LEFT JOIN TPRODUCTO p ON d.tCodigoProducto = p.tCodigoProducto
            WHERE d.tCodigoPedido = @CodigoPedido
              AND ISNULL(d.tEstadoItem, 'N') = 'N'
            ORDER BY d.tItem
            """;

        var rows = await connection.QueryAsync<ItemPedidoVistaRow>(sql, new { CodigoPedido = codigoPedido });
        return rows.Select(r => new ItemPedidoVista(
            Item: r.Item,
            CodigoProducto: r.CodigoProducto,
            Producto: r.Producto,
            PrecioOficial: r.PrecioOficial,
            Descuento: r.Descuento,
            PrecioVenta: r.PrecioVenta,
            Cantidad: r.Cantidad,
            SubTotal: r.SubTotal,
            EstadoItem: r.EstadoItem,
            Facturado: r.Facturado == "S",
            Imprime: r.Imprime,
            TienePropiedad: r.TienePropiedad,
            TieneObservacion: r.TieneObservacion,
            Corte: r.Corte,
            Observacion: r.Observacion,
            Documento: r.Documento,
            Usuario: r.Usuario,
            FechaEnvio: r.FechaEnvio,
            Comanda: r.Comanda))
        .ToList()
        .AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<ComboPedidoVista>> ObtenerCombosAsync(
        string codigoPedido, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // Legacy: vPedidoCombo — CPEDIDO JOIN TPRODUCTO.
        const string sql = """
            SELECT
                c.tItem                     AS Item,
                c.tItemCombo                AS ItemCombo,
                c.tProducto                 AS CodigoProducto,
                c.tProductoCombo            AS CodigoProductoCombo,
                ISNULL(p.tDetallado, c.tProductoCombo) AS Producto,
                CAST(ISNULL(c.nCantidad, 0) AS decimal(18,4)) AS Cantidad,
                CAST(ISNULL(c.lImprime, 0)  AS bit)           AS Imprime,
                CAST(CASE WHEN DATALENGTH(c.tObservacion) > 0 THEN 1 ELSE 0 END AS bit) AS TieneObservacion,
                CAST(0 AS bit)              AS TienePropiedad,
                c.tObservacion              AS Observacion
            FROM CPEDIDO c
            LEFT JOIN TPRODUCTO p ON c.tProductoCombo = p.tCodigoProducto
            WHERE c.tCodigoPedido = @CodigoPedido
            ORDER BY c.tItem, c.tItemCombo
            """;

        var rows = await connection.QueryAsync<ComboPedidoVistaRow>(sql, new { CodigoPedido = codigoPedido });
        return rows.Select(r => new ComboPedidoVista(
            Item: r.Item,
            ItemCombo: r.ItemCombo,
            CodigoProducto: r.CodigoProducto,
            CodigoProductoCombo: r.CodigoProductoCombo,
            Producto: r.Producto,
            Cantidad: r.Cantidad,
            Imprime: r.Imprime,
            TieneObservacion: r.TieneObservacion,
            TienePropiedad: r.TienePropiedad,
            Observacion: r.Observacion))
        .ToList()
        .AsReadOnly();
    }

    /// <inheritdoc/>
    public async Task<IReadOnlyList<PedidoSinMesaVista>> ObtenerActivosSinMesaAsync(
        string caja, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // Legacy: frmMesas.frm — pedidos activos sin mesa para el panel lateral.
        // BR-MESAS-005: tEstadoPedido='01', tTipoPedido<>'04', mesa vacía, ordenado desc por pedido.
        const string sql = """
            SELECT tCodigoPedido AS CodigoPedido,
                   ISNULL(tObservacion, '') AS Observacion
            FROM MPEDIDO
            WHERE tEstadoPedido = '01'
              AND tTipoPedido   <> '04'
              AND LEN(RTRIM(ISNULL(tMesa, ''))) = 0
              AND tCaja = @Caja
            ORDER BY tCodigoPedido DESC
            """;
        var rows = await connection.QueryAsync(sql, new { Caja = caja });
        return rows.Select(r => new PedidoSinMesaVista(
            CodigoPedido: (string)r.CodigoPedido,
            Observacion: (string)r.Observacion))
            .ToList().AsReadOnly();
    }

    // ── POS-FUNC-034 ──────────────────────────────────────────────────────────

    /// <inheritdoc/>
    public async Task<IReadOnlyList<PedidoCorrelativoVista>> ObtenerCorrelativoAsync(
        DateOnly desde, DateOnly hasta, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // Legacy: frmPedidoCorrelativo.frm — vPedidoCorrelativo WHERE fFecha BETWEEN @desde AND @hasta.
        // BR-CORRPEDIDO-001: rango de fechas de día completo (00:00 - 23:59).
        const string sql = """
            SELECT tCodigoPedido      AS CodigoPedido,
                   ISNULL(tCaja,'')   AS Caja,
                   ISNULL(tTurno,'')  AS Turno,
                   fFecha             AS Fecha,
                   ISNULL(Mesa,'')    AS Mesa,
                   ISNULL(Mozo,'')    AS Mozo,
                   ISNULL(nVenta,0)   AS Venta,
                   ISNULL(TipoPedido,'') AS TipoPedido,
                   ISNULL(Estado,'')  AS Estado,
                   ISNULL(Documento,'') AS Documento,
                   ISNULL(tObservacion,'') AS Observacion,
                   ISNULL(nAdulto,0)  AS Adultos
            FROM vPedidoCorrelativo
            WHERE fFecha >= @Desde AND fFecha < @HastaExclusivo
            ORDER BY tCodigoPedido
            """;
        var rows = await connection.QueryAsync(sql, new
        {
            Desde = desde.ToDateTime(TimeOnly.MinValue),
            HastaExclusivo = hasta.ToDateTime(TimeOnly.MinValue).AddDays(1)
        });
        return rows.Select(r => new PedidoCorrelativoVista(
            CodigoPedido: (string)r.CodigoPedido,
            Caja: (string)r.Caja,
            Turno: (string)r.Turno,
            Fecha: (DateTime)r.Fecha,
            Mesa: (string)r.Mesa,
            Mozo: (string)r.Mozo,
            Venta: (decimal)r.Venta,
            TipoPedido: (string)r.TipoPedido,
            Estado: (string)r.Estado,
            Documento: (string)r.Documento,
            Observacion: (string)r.Observacion,
            Adultos: (decimal)r.Adultos))
            .ToList().AsReadOnly();
    }

    /// <inheritdoc/>
    public async Task<IReadOnlyList<DocumentoAgrupadoVista>> ObtenerDocumentosAgrupadosPedidoAsync(
        string codigoPedido, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // Legacy: frmPedidoAnterior.frm — vDocumentoAgrupado WHERE tCodigoPedido = @codigoPedido.
        const string sql = """
            SELECT tDocumento                      AS Documento,
                   ISNULL(nVenta,0)                AS Venta,
                   ISNULL(Estado,'')               AS Estado,
                   ISNULL(tTurno,'')               AS Turno,
                   ISNULL(tObservacion,'')         AS ObservacionAnulacion,
                   ISNULL(tUsuarioAnulado,'')      AS UsuarioAnulado,
                   fRegistroAnulado                AS FechaAnulacion
            FROM vDocumentoAgrupado
            WHERE tCodigoPedido = @CodigoPedido
            ORDER BY tDocumento
            """;
        var rows = await connection.QueryAsync(sql, new { CodigoPedido = codigoPedido });
        return rows.Select(r => new DocumentoAgrupadoVista(
            Documento: (string)r.Documento,
            Venta: (decimal)r.Venta,
            Estado: (string)r.Estado,
            Turno: (string)r.Turno,
            ObservacionAnulacion: (string)r.ObservacionAnulacion,
            UsuarioAnulado: (string)r.UsuarioAnulado,
            FechaAnulacion: (DateTime?)r.FechaAnulacion))
            .ToList().AsReadOnly();
    }
}
