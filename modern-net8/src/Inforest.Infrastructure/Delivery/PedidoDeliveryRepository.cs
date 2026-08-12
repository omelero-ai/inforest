using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Infrastructure.Delivery;

/// <summary>
/// Repositorio Dapper para pedidos delivery.
/// <para>
/// Legacy: tablas <c>MPEDIDO</c>/<c>DPEDIDO</c>.
/// Vistas: <c>vDespachador</c>, <c>vDelivery</c>, <c>vPedidoCabecera</c>,
///   <c>vPedidoDetalle</c>, <c>vPedidoGrilla</c>.
/// SPs: <c>spUpdate_DPEDIDO_Ina</c> — inactivación de ítem de detalle.
/// </para>
/// </summary>
internal sealed class PedidoDeliveryRepository : IPedidoDeliveryRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public PedidoDeliveryRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    /// <inheritdoc />
    public async Task<PedidoDelivery?> ObtenerPorCodigoAsync(string codigoPedido, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoPedido AS CodigoPedido, tClienteDelivery AS CodigoClienteDelivery,
                   tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   fFecha AS FechaRegistro, nMonto AS MontoTotal,
                   tEstadoPedido AS EstadoDelivery, tTipoPedido AS TipoPedido,
                   lPrioridad AS EsPrioridad
            FROM MPEDIDO
            WHERE tCodigoPedido = @CodigoPedido
            """;
        return await conn.QueryFirstOrDefaultAsync<PedidoDelivery>(sql, new { CodigoPedido = codigoPedido });
    }

    /// <inheritdoc />
    public async Task<IEnumerable<PedidoDelivery>> ObtenerPendientesAsync(
        string codigoCaja, DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoPedido AS CodigoPedido, tClienteDelivery AS CodigoClienteDelivery,
                   tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   fFecha AS FechaRegistro, nMonto AS MontoTotal,
                   tEstadoPedido AS EstadoDelivery, tTipoPedido AS TipoPedido,
                   lPrioridad AS EsPrioridad
            FROM vDelivery
            WHERE tCaja = @Caja
              AND fFecha >= @FechaInicio AND fFecha <= @FechaFin
              AND tEstadoPedido NOT IN ('03','04')
            ORDER BY fFecha ASC
            """;
        return await conn.QueryAsync<PedidoDelivery>(sql,
            new { Caja = codigoCaja, FechaInicio = fechaInicio, FechaFin = fechaFin });
    }

    /// <inheritdoc />
    public async Task<IEnumerable<PedidoDelivery>> ObtenerPorEstadoAsync(
        EstadoDelivery estado, DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoPedido AS CodigoPedido, tClienteDelivery AS CodigoClienteDelivery,
                   tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   fFecha AS FechaRegistro, nMonto AS MontoTotal,
                   tEstadoPedido AS EstadoDelivery, tTipoPedido AS TipoPedido
            FROM vDelivery
            WHERE fFecha >= @FechaInicio AND fFecha <= @FechaFin
              AND tEstadoPedido = @Estado
            ORDER BY fFecha ASC
            """;
        return await conn.QueryAsync<PedidoDelivery>(sql,
            new { FechaInicio = fechaInicio, FechaFin = fechaFin, Estado = ((int)estado).ToString("D2") });
    }

    /// <inheritdoc />
    public async Task<IEnumerable<PedidoDelivery>> ObtenerParaDespachadorAsync(
        DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        // Legacy: vDespachador where tTipoPedido='02' and tEstadoPedido='02' and fLlegada IS NULL
        const string sql = """
            SELECT tCodigoPedido AS CodigoPedido, tClienteDelivery AS CodigoClienteDelivery,
                   tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   fFecha AS FechaRegistro, nMonto AS MontoTotal,
                   tEstadoPedido AS EstadoDelivery, tTipoPedido AS TipoPedido,
                   lPrioridad AS EsPrioridad, fDespacho AS FechaDespacho
            FROM vDespachador
            WHERE tTipoPedido = '02' AND tEstadoPedido = '02'
              AND ISNULL(fLlegada, 0) = 0
              AND fFecha >= @FechaInicio AND fFecha <= @FechaFin
            ORDER BY lEmpacador, fFecha ASC
            """;
        return await conn.QueryAsync<PedidoDelivery>(sql,
            new { FechaInicio = fechaInicio, FechaFin = fechaFin });
    }

    /// <inheritdoc />
    public async Task InsertarAsync(PedidoDelivery pedido, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            INSERT INTO MPEDIDO
                (tCodigoPedido, tClienteDelivery, tZona, tDistrito,
                 fFecha, nMonto, tEstadoPedido, tTipoPedido, lPrioridad,
                 tTurno, tCaja, tMozo, fHoraEstimadaEntrega)
            VALUES
                (@CodigoPedido, @CodigoClienteDelivery, @CodigoZona, @CodigoDistrito,
                 @FechaRegistro, @MontoTotal, '00', @TipoPedido, @EsPrioridad,
                 @CodigoTurno, @CodigoCaja, @CodigoUsuario, @HoraEstimadaEntrega)
            """;
        await conn.ExecuteAsync(sql, pedido);
    }

    /// <inheritdoc />
    public async Task ActualizarEstadoAsync(string codigoPedido, EstadoDelivery estado, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = "UPDATE MPEDIDO SET tEstadoPedido = @Estado WHERE tCodigoPedido = @CodigoPedido";
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido, Estado = ((int)estado).ToString("D2") });
    }

    /// <inheritdoc />
    public async Task AsignarMotorizadoAsync(string codigoPedido, string codigoMotorizado, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = "UPDATE MPEDIDO SET tMotorizado = @Motorizado WHERE tCodigoPedido = @CodigoPedido";
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido, Motorizado = codigoMotorizado });
    }
}
