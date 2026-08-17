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
    public async Task<IEnumerable<PedidoReasignacionMotorizado>> ObtenerPedidosReasignacionAsync(
        DateTime fechaInicio, DateTime fechaFin, string? criterioPedido, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT m.tCodigoPedido AS CodigoPedido,
                   mz.Descripcion AS MotorizadoDescripcion,
                   CAST(m.nTarifaMotorizadoN AS decimal(18,2)) AS TarifaMotorizado
            FROM MPEDIDO m
            LEFT JOIN vMotorizado mz ON m.tMotorizadoN = mz.Codigo
            WHERE m.tTipoPedido = '02'
              AND m.tEstadoPedido = '02'
              AND ISNULL(m.tMotorizado, '') <> '0000'
              AND m.fRegistro >= @FechaInicio
              AND m.fRegistro <= @FechaFin
              AND (@CriterioPedido = '' OR m.tCodigoPedido LIKE '%' + @CriterioPedido)
            ORDER BY m.fRegistro, m.tCodigoPedido
            """;
        return await conn.QueryAsync<PedidoReasignacionMotorizado>(
            sql,
            new { FechaInicio = fechaInicio, FechaFin = fechaFin, CriterioPedido = criterioPedido ?? string.Empty });
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
    public async Task<IReadOnlyList<PedidoDeliverySeguimiento>> ObtenerSeguimientoEntregadosAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoPedido AS CodigoPedido,
                   fFecha AS FechaRegistro,
                   tUsuario AS Usuario,
                   tCaja AS Caja,
                   tTelefono AS Telefono,
                   Cliente AS Cliente,
                   Empacador AS Empacador,
                   Motorizado AS Motorizado,
                   fAsignacion AS FechaAsignacion,
                   fSalida AS FechaSalida,
                   fLlegada AS FechaLlegada,
                   Referencia AS Referencia,
                   tDireccion AS Direccion,
                   Zona AS Zona
            FROM vDespachador
            WHERE tTipoPedido = '02'
              AND tEstadoPedido = '02'
              AND ISNULL(fLlegada, 0) <> 0
              AND fFecha >= @FechaInicio
              AND fFecha <= @FechaFin
            ORDER BY tCodigoPedido
            """;
        var rows = await conn.QueryAsync<PedidoDeliverySeguimiento>(sql, new
        {
            FechaInicio = fechaInicio,
            FechaFin = fechaFin
        });
        return rows.ToList();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<PedidoDespachadorResumen>> ObtenerResumenDespachadorAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoPedido AS CodigoPedido,
                   fFecha AS FechaRegistro,
                   tUsuario AS Usuario,
                   tCaja AS Caja,
                   Cliente AS Cliente,
                   tTelefono AS Telefono,
                   tDireccion AS Direccion,
                   Referencia AS Referencia,
                   Empacador AS Empacador,
                   Motorizado AS Motorizado,
                   fAsignacion AS FechaAsignacion,
                   CAST(ISNULL(nMonto, 0) AS decimal(18,2)) AS MontoTotal
            FROM vDespachador
            WHERE tTipoPedido IN (SELECT Codigo FROM vTipoPedido WHERE lActivo = 1 AND lCanalDelivery = 1)
              AND tEstadoPedido = '02'
              AND ISNULL(fLlegada, 0) = 0
              AND fFecha >= @FechaInicio
              AND fFecha <= @FechaFin
            ORDER BY lEmpacador, fFecha ASC
            """;
        var rows = await conn.QueryAsync<PedidoDespachadorResumen>(sql, new
        {
            FechaInicio = fechaInicio,
            FechaFin = fechaFin
        });
        return rows.ToList();
    }

    /// <inheritdoc />
    public async Task<PedidoDespachadorResumen?> ObtenerResumenDespachadorPorPedidoAsync(
        string codigoPedido,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT TOP 1 tCodigoPedido AS CodigoPedido,
                   fFecha AS FechaRegistro,
                   tUsuario AS Usuario,
                   tCaja AS Caja,
                   Cliente AS Cliente,
                   tTelefono AS Telefono,
                   tDireccion AS Direccion,
                   Referencia AS Referencia,
                   Empacador AS Empacador,
                   Motorizado AS Motorizado,
                   fAsignacion AS FechaAsignacion,
                   CAST(ISNULL(nMonto, 0) AS decimal(18,2)) AS MontoTotal
            FROM vDespachador
            WHERE tCodigoPedido = @CodigoPedido
            """;
        return await conn.QueryFirstOrDefaultAsync<PedidoDespachadorResumen>(sql, new { CodigoPedido = codigoPedido });
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<OperadorDespachoItem>> ObtenerMotorizadosActivosDespachoAsync(
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT Codigo, Descripcion
            FROM vMotorizado
            WHERE lActivo = 1
              AND Codigo <> '0000'
            ORDER BY Descripcion
            """;
        var rows = await conn.QueryAsync<OperadorDespachoItem>(sql);
        return rows.ToList();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<OperadorDespachoItem>> ObtenerEmpacadoresActivosDespachoAsync(
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT Codigo, Descripcion
            FROM vEmpacador
            WHERE lActivo = 1
            ORDER BY Descripcion
            """;
        var rows = await conn.QueryAsync<OperadorDespachoItem>(sql);
        return rows.ToList();
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

    /// <inheritdoc />
    public async Task AsignarMotorizadoDespachoAsync(
        string codigoPedido,
        string codigoMotorizado,
        decimal tarifaDiaria,
        bool esTarifaExtra,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        using var tx = conn.BeginTransaction();

        const string sqlMarcaMotorizado = """
            UPDATE TTABLA
            SET ntamano = 1,
                lreplica = 1
            WHERE tTABLA = 'MOTORIZADO'
              AND tCodigo = @CodigoMotorizado
            """;

        const string sqlPedido = """
            UPDATE MPEDIDO
            SET nTarifaMotorizado = @TarifaDiaria,
                nTarifaExtra = @TarifaExtra,
                fSalida = GETDATE(),
                fAsignacion = GETDATE(),
                tMotorizado = @CodigoMotorizado
            WHERE tCodigoPedido = @CodigoPedido
            """;

        try
        {
            await conn.ExecuteAsync(sqlMarcaMotorizado, new { CodigoMotorizado = codigoMotorizado }, tx);
            await conn.ExecuteAsync(sqlPedido, new
            {
                CodigoPedido = codigoPedido,
                CodigoMotorizado = codigoMotorizado,
                TarifaDiaria = tarifaDiaria,
                TarifaExtra = esTarifaExtra ? 1 : 0
            }, tx);

            tx.Commit();
        }
        catch
        {
            tx.Rollback();
            throw;
        }
    }

    /// <inheritdoc />
    public async Task DesasignarMotorizadoDespachoAsync(
        string codigoPedido,
        string codigoMotorizado,
        bool liberarBanderaMotorizado,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        using var tx = conn.BeginTransaction();

        try
        {
            if (liberarBanderaMotorizado)
            {
                const string sqlTabla = """
                    UPDATE TTABLA
                    SET ntamano = 0,
                        lreplica = 1
                    WHERE tTABLA = 'MOTORIZADO'
                      AND tCodigo = @CodigoMotorizado
                    """;
                await conn.ExecuteAsync(sqlTabla, new { CodigoMotorizado = codigoMotorizado }, tx);
            }

            const string sqlPedido = """
                UPDATE MPEDIDO
                SET nTarifaMotorizado = NULL,
                    nTarifaExtra = NULL,
                    fAsignacion = NULL,
                    fSalida = NULL,
                    tMotorizado = '0000'
                WHERE tCodigoPedido = @CodigoPedido
                """;
            await conn.ExecuteAsync(sqlPedido, new { CodigoPedido = codigoPedido }, tx);

            tx.Commit();
        }
        catch
        {
            tx.Rollback();
            throw;
        }
    }

    /// <inheritdoc />
    public async Task AsignarEmpacadorDespachoAsync(
        string codigoPedido,
        string codigoEmpacador,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MPEDIDO
            SET tEmpacador = @CodigoEmpacador,
                fEmpacador = GETDATE()
            WHERE tCodigoPedido = @CodigoPedido
            """;
        await conn.ExecuteAsync(sql, new
        {
            CodigoPedido = codigoPedido,
            CodigoEmpacador = codigoEmpacador
        });
    }

    /// <inheritdoc />
    public async Task DesasignarEmpacadorDespachoAsync(string codigoPedido, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MPEDIDO
            SET tEmpacador = '',
                fEmpacador = NULL
            WHERE tCodigoPedido = @CodigoPedido
            """;
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido });
    }

    /// <inheritdoc />
    public async Task<int> ContarPedidosActivosMotorizadoAsync(
        string codigoMotorizado,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT COUNT(1)
            FROM vDespachador
            WHERE tTipoPedido = '02'
              AND tEstadoPedido = '02'
              AND ISNULL(fLlegada, 0) = 0
              AND tMotorizado = @CodigoMotorizado
            """;
        return await conn.ExecuteScalarAsync<int>(sql, new { CodigoMotorizado = codigoMotorizado });
    }

    /// <inheritdoc />
    public async Task<int> ContarAsignacionesPrincipalesMotorizadoAsync(
        string codigoMotorizado,
        DateTime fechaOperacion,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT COUNT(ISNULL(nTarifaMotorizado, 0))
            FROM MPEDIDO
            WHERE tMotorizado = @CodigoMotorizado
              AND ISNULL(nTarifaExtra, 0) = 0
              AND CONVERT(nvarchar(8), fAsignacion, 112) = @FechaOperacion
            """;
        return await conn.ExecuteScalarAsync<int>(
            sql,
            new { CodigoMotorizado = codigoMotorizado, FechaOperacion = fechaOperacion.ToString("yyyyMMdd") });
    }

    /// <inheritdoc />
    public async Task<int> ContarReasignacionesMotorizadoAsync(
        string codigoMotorizado,
        DateTime fechaOperacion,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT COUNT(ISNULL(nTarifaMotorizado, 0))
            FROM MPEDIDO
            WHERE tMotorizadoN = @CodigoMotorizado
              AND ISNULL(nTarifaExtraN, 0) = 0
              AND CONVERT(nvarchar(8), fAsignacion, 112) = @FechaOperacion
            """;
        return await conn.ExecuteScalarAsync<int>(
            sql,
            new { CodigoMotorizado = codigoMotorizado, FechaOperacion = fechaOperacion.ToString("yyyyMMdd") });
    }

    /// <inheritdoc />
    public async Task ActualizarReasignacionMotorizadoAsync(
        string codigoPedido,
        string codigoMotorizado,
        decimal tarifaMotorizado,
        bool esTarifaExtra,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MPEDIDO
            SET nTarifaMotorizadoN = @TarifaMotorizado,
                tMotorizadoN = @CodigoMotorizado,
                nTarifaExtraN = @EsTarifaExtra
            WHERE tCodigoPedido = @CodigoPedido
            """;
        await conn.ExecuteAsync(sql, new
        {
            CodigoPedido = codigoPedido,
            CodigoMotorizado = codigoMotorizado,
            TarifaMotorizado = tarifaMotorizado,
            EsTarifaExtra = esTarifaExtra ? 1 : 0
        });
    }

    /// <inheritdoc />
    public async Task LimpiarReasignacionMotorizadoAsync(string codigoPedido, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MPEDIDO
            SET nTarifaMotorizadoN = NULL,
                nTarifaExtraN = NULL,
                tMotorizadoN = NULL
            WHERE tCodigoPedido = @CodigoPedido
            """;
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido });
    }

    // ── CentralPedidos / frmCentralPedidos.frm ────────────────────────────────

    /// <inheritdoc />
    public async Task ConfirmarEntregaAsync(string codigoPedido, string usuario, CancellationToken ct = default)
    {
        // Legacy: Update MPEDIDO Set lEntregado='1', tusuarioentregado=@sUsuario, fregentregado=getdate()
        //         frmCentralPedidos.frm Case 3 — BR-DEL-012
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MPEDIDO
            SET lEntregado = 1,
                tusuarioentregado = @Usuario,
                fregentregado = GETDATE()
            WHERE tCodigoPedido = @CodigoPedido
            """;
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido, Usuario = usuario });
    }

    /// <inheritdoc />
    public async Task RevertirEntregaAsync(string codigoPedido, string usuario, CancellationToken ct = default)
    {
        // Legacy: Update MPEDIDO Set lEntregado='0', tusuarioentregado=@sUsuario, fregentregado=getdate()
        //         frmCentralPedidos.frm Case 5 — BR-DEL-012
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MPEDIDO
            SET lEntregado = 0,
                tusuarioentregado = @Usuario,
                fregentregado = GETDATE()
            WHERE tCodigoPedido = @CodigoPedido
            """;
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido, Usuario = usuario });
    }

    /// <inheritdoc />
    public async Task ModificarFechaProgramadaAsync(string codigoPedido, DateTime nuevaFecha, CancellationToken ct = default)
    {
        // Legacy: Update MPEDIDO set fregistro=@fecha, fProgramacion=@fecha where tCodigoPedido=@sPedido
        //         frmCentralPedidos.frm Case 2 — BR-DEL-014
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE MPEDIDO
            SET fregistro = @NuevaFecha,
                fProgramacion = @NuevaFecha
            WHERE tCodigoPedido = @CodigoPedido
            """;
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido, NuevaFecha = nuevaFecha });
    }

    /// <inheritdoc />
    public async Task<string?> ObtenerEstadoPagoAsync(string codigoPedido, CancellationToken ct = default)
    {
        // Legacy: Select * From vDocumentoAgrupado Where tCodigoPedido=@sPedido
        //         frmCentralPedidos.frm Case 3 — detecta 'POR COBRAR' / 'ANTICIPO' — BR-DEL-013
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT TOP 1 ISNULL(Estado, 'PAGADO')
            FROM vDocumentoAgrupado
            WHERE tCodigoPedido = @CodigoPedido
            """;
        return await conn.ExecuteScalarAsync<string?>(sql, new { CodigoPedido = codigoPedido });
    }

    /// <inheritdoc />
    public async Task<bool> EstaEntregadoAsync(string codigoPedido, CancellationToken ct = default)
    {
        // Legacy: grdGrilla.Columns(6).Text = "ENTREGADO" from MPEDIDO.lEntregado
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = "SELECT ISNULL(lEntregado, 0) FROM MPEDIDO WHERE tCodigoPedido = @CodigoPedido";
        return await conn.ExecuteScalarAsync<bool>(sql, new { CodigoPedido = codigoPedido });
    }
}
