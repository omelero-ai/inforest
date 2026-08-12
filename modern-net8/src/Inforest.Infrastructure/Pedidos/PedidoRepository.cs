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
}
