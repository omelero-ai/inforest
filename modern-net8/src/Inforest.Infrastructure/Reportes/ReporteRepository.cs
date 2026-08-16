using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Reportes;

/// <summary>
/// Implementación de <see cref="IReporteRepository"/> usando Dapper sobre SQL Server.
/// <para>
/// Todos los SPs se invocan por nombre exacto del Legacy (ADR-002, ADR-007).
/// El timeout es de 600 segundos (BR-SQL-CMD-001).
/// Los parámetros <c>@sPrecio</c> se construyen desde <see cref="ExpresionPrecioExtensions"/>
/// para evitar SQL injection (BR-REP-SQL-DYN-001).
/// </para>
/// Legacy: 5. SP.sql — spRep_*, USP_KDS_Resporte*
/// Reglas: BR-REP-001 … BR-REP-012
/// </summary>
internal sealed class ReporteRepository : IReporteRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;
    private readonly ILogger<ReporteRepository> _logger;

    public ReporteRepository(
        IDbConnectionFactory connectionFactory,
        ISpExecutor spExecutor,
        ILogger<ReporteRepository> logger)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
        _logger = logger;
    }

    // ── Grupo A — Core Operativo ──────────────────────────────────────────────

    /// <inheritdoc />
    public async Task<IReadOnlyList<ComandaRow>> ObtenerComandaAsync(
        bool flagTipo,
        string orden,
        DateTime fechaInicio,
        DateTime fechaFin,
        string criterio,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte Comanda: {FechaInicio} – {FechaFin}", fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<ComandaRow>(
            conn,
            "spRep_Comanda",
            new
            {
                flagTipo,
                SOrden = orden,
                fInicio = fechaInicio,
                fFinal = fechaFin,
                sCriterio = criterio
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<PropinaRow>> ObtenerPropinaAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        string condicion,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte Propina: {FechaInicio} – {FechaFin}", fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<PropinaRow>(
            conn,
            "spRep_Propina",
            new
            {
                fInicio = fechaInicio,
                fFinal = fechaFin,
                sCondicion = condicion
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<PrincipalClienteDetalleRow>> ObtenerPrincipalClienteDetalleAsync(
        double montoMinimo,
        string codigoCliente,
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte PrincipalCliente detalle: {FechaInicio} – {FechaFin}", fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<PrincipalClienteDetalleRow>(
            conn,
            "spRep_PrincipalCliente",
            new
            {
                flagTipo = true,
                sMonto = montoMinimo,
                SCliente = codigoCliente,
                fInicio = fechaInicio,
                fFinal = fechaFin
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<PrincipalClienteResumenRow>> ObtenerPrincipalClienteResumenAsync(
        double montoMinimo,
        string codigoCliente,
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte PrincipalCliente resumen: {FechaInicio} – {FechaFin}", fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<PrincipalClienteResumenRow>(
            conn,
            "spRep_PrincipalCliente",
            new
            {
                flagTipo = false,
                sMonto = montoMinimo,
                SCliente = codigoCliente,
                fInicio = fechaInicio,
                fFinal = fechaFin
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<PaloteoComparativoRow>> ObtenerPaloteoComparativoAsync(
        bool porTurno,
        bool porValor,
        string tipoProducto,
        string areaProduccion,
        string grupo,
        string subGrupo,
        string codigoProducto,
        string turno,
        DateTime fechaInicio,
        DateTime fechaFin,
        bool soloFacturados,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte PaloteoComparativo: {FechaInicio} – {FechaFin}", fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<PaloteoComparativoRow>(
            conn,
            "spRep_PaloteoComparativo",
            new
            {
                flagTurnoOFecha = porTurno,
                flagTipoValor = porValor,
                tTipoProducto = tipoProducto,
                tAreaProduccion = areaProduccion,
                tGrupo = grupo,
                tSubGrupo = subGrupo,
                tCodigoProducto = codigoProducto,
                tTurno = turno,
                finicio = fechaInicio,
                ffinal = fechaFin,
                flagNFacturado = soloFacturados
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<PaloteoSubProdRow>> ObtenerPaloteoSubProdAsync(
        bool porTurno,
        string turno,
        string local,
        string salon,
        string tipoProducto,
        string mozo,
        string tipoPedido,
        string area,
        string caja,
        string codigoProducto,
        string codigoCliente,
        string boton2,
        string boton3,
        string boton4,
        string boton5,
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte PaloteoSubProd: {FechaInicio} – {FechaFin}", fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<PaloteoSubProdRow>(
            conn,
            "spRep_PaloteoSubProd",
            new
            {
                flagTurno = porTurno,
                tTurno = turno,
                tLocal = local,
                tSalon = salon,
                tipoProd = tipoProducto,
                tmozo = mozo,
                ttipoPedido = tipoPedido,
                tarea = area,
                tcaja = caja,
                tcodigoproducto = codigoProducto,
                tcodigocliente = codigoCliente,
                sBoton2 = boton2,
                sBoton3 = boton3,
                sBoton4 = boton4,
                sBoton5 = boton5,
                fInicio = fechaInicio,
                fFinal = fechaFin
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    // ── Grupo B — Integrados Multi-Local ─────────────────────────────────────

    /// <inheritdoc />
    public async Task<IReadOnlyList<CtaCteIntegradoRow>> ObtenerCtaCteIntegradoAsync(
        string flagTipo,
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte CtaCteIntegrado tipo={Tipo}: {FechaInicio} – {FechaFin}", flagTipo, fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<CtaCteIntegradoRow>(
            conn,
            "spRep_CtaCteIntegrado",
            new
            {
                flagTipo,
                fInicio = fechaInicio,
                fFinal = fechaFin
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<CtaCteOperativaRow>> ObtenerCtaCteOperativaAsync(
        CtaCteOperativaParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte CtaCte operativa: {FechaInicio} – {FechaFin}", p.FechaInicio, p.FechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<CtaCteOperativaRow>(
            conn,
            "spRep_CtaCteN",
            new
            {
                flagTDetalle = p.FlagDetalle,
                flagTResumido = p.FlagResumido,
                flagTConsolidado = p.FlagConsolidado,
                fInicio = p.FechaInicio,
                fFinal = p.FechaFin,
                Estado = p.Estado,
                Cliente = p.Cliente,
                TipoCC = p.TipoCtaCte,
                SubTipoCC = p.SubTipoCtaCte
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<ReporteFiltroOpcion>> ObtenerTiposCtaCteAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT Codigo, Descripcion
            FROM dbo.vTipoCtaCte
            WHERE lActivo = 1
            ORDER BY Descripcion
            """;
        var result = await conn.QueryAsync<ReporteFiltroOpcion>(new CommandDefinition(sql, cancellationToken: ct));
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<ReporteFiltroOpcion>> ObtenerSubTiposCtaCteAsync(
        string tipoCtaCte,
        CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT Codigo, Descripcion
            FROM dbo.vSubTipoCtaCte
            WHERE lActivo = 1
              AND (@TipoCtaCte = '' OR tTipoCtaCte = @TipoCtaCte)
            ORDER BY Descripcion
            """;
        var result = await conn.QueryAsync<ReporteFiltroOpcion>(
            new CommandDefinition(sql, new { TipoCtaCte = tipoCtaCte }, cancellationToken: ct));
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<PaloteoVentaIntegradoRow>> ObtenerPaloteoVentaIntegradoAsync(
        PaloteoVentaIntegradoParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte PaloteoVentaIntegrado: {FechaInicio} – {FechaFin}", p.FechaInicio, p.FechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        // BR-REP-SQL-DYN-001: expresión precio desde enum, nunca string libre
        var sPrecio = p.TipoPrecio.ToSqlExpresion();
        var result = await _spExecutor.QueryAsync<PaloteoVentaIntegradoRow>(
            conn,
            "spRep_PaloteoVentaIntegrado",
            new
            {
                flagPropiedades = p.FlagPropiedades,
                flagProduccion = p.FlagProduccion,
                flagVenta = p.FlagVenta,
                flagCortesia = p.FlagCortesia,
                flagCuentaCte = p.FlagCuentaCte,
                flagPedidosFacturados = p.FlagPedidosFacturados,
                flagCombinacion = p.FlagCombinacion,
                flagCargo = p.FlagCargo,
                flagOpcion = p.FlagOpcion,
                flagVVenta = p.TipoPrecio == ExpresionPrecio.Venta,
                flagVNeto = p.TipoPrecio == ExpresionPrecio.Neto,
                flagVCosto = p.TipoPrecio == ExpresionPrecio.Costo,
                tTipoProducto = p.TipoProducto,
                tTipoPedido = p.TipoPedido,
                tGrupo = p.Grupo,
                tSubGrupo = p.SubGrupo,
                sPrecio,
                sNeto = ExpresionPrecio.Neto.ToSqlExpresion(),
                sCosto = ExpresionPrecio.Costo.ToSqlExpresion(),
                sOrden = p.Orden,
                sBoton2 = p.Boton2,
                sBoton3 = p.Boton3,
                sBoton4 = p.Boton4,
                sBoton5 = p.Boton5,
                finicio = p.FechaInicio,
                ffinal = p.FechaFin
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<RankingIntegradoRow>> ObtenerRankingIntegradoAsync(
        RankingIntegradoParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte RankingIntegrado: {FechaInicio} – {FechaFin}", p.FechaInicio, p.FechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var sPrecio = p.TipoPrecio.ToSqlExpresion();
        var result = await _spExecutor.QueryAsync<RankingIntegradoRow>(
            conn,
            "spRep_RankingIntegrado",
            new
            {
                flagProduccion = p.FlagProduccion,
                flagVenta = p.FlagVenta,
                flagCortesia = p.FlagCortesia,
                flagCuentaCte = p.FlagCuentaCte,
                flagPedidosFacturados = p.FlagPedidosFacturados,
                flagCombinacion = p.FlagCombinacion,
                flagCargo = p.FlagCargo,
                flagOpcion = p.FlagOpcion,
                flagVVenta = p.TipoPrecio == ExpresionPrecio.Venta,
                flagVNeto = p.TipoPrecio == ExpresionPrecio.Neto,
                flagVCosto = p.TipoPrecio == ExpresionPrecio.Costo,
                tTipoProducto = p.TipoProducto,
                tTipoPedido = p.TipoPedido,
                tGrupo = p.Grupo,
                tSubGrupo = p.SubGrupo,
                sPrecio,
                sOrden = p.Orden,
                sBoton2 = p.Boton2,
                sBoton3 = p.Boton3,
                sBoton4 = p.Boton4,
                sBoton5 = p.Boton5,
                sNeto = ExpresionPrecio.Neto.ToSqlExpresion(),
                sCosto = ExpresionPrecio.Costo.ToSqlExpresion(),
                finicio = p.FechaInicio,
                ffinal = p.FechaFin
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<VentaMensualIntegradoRow>> ObtenerVentaMensualIntegradoAsync(
        VentaMensualIntegradoParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte VentaMensualIntegrado: {Ano}/{Mes}", p.Ano, p.Mes);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        // BR-REP-SQL-DYN-001
        var sPrecio = p.TipoPrecio.ToSqlExpresion();
        var result = await _spExecutor.QueryAsync<VentaMensualIntegradoRow>(
            conn,
            "spRep_VentaMensualIntegrado",
            new
            {
                sPrecio,
                sAno = p.Ano,
                sMes = p.Mes,
                sFecha = p.CondicionFecha,
                dHour = p.HoraCorte,
                tipo = p.Tipo,
                metaMensual = p.MetaMensual,
                diames = p.DiasDelMes
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    // ── Grupo C — Delivery / Motorizado ──────────────────────────────────────

    /// <inheritdoc />
    public async Task<IReadOnlyList<AnaliticoMotorizadoIntegradoRow>> ObtenerAnaliticoMotorizadoIntegradoAsync(
        AnaliticoMotorizadoIntegradoParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte AnaliticoMotorizadoIntegrado: {FechaInicio} – {FechaFin}", p.FechaInicio, p.FechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var sPrecio = p.TipoPrecio.ToSqlExpresion();
        var result = await _spExecutor.QueryAsync<AnaliticoMotorizadoIntegradoRow>(
            conn,
            "spRep_AnaliticoMotorizadoIntegrado",
            new
            {
                flagProduccion = p.FlagProduccion,
                flagVenta = p.FlagVenta,
                flagCortesia = p.FlagCortesia,
                flagCuentaCte = p.FlagCuentaCte,
                flagPedidosFacturados = p.FlagPedidosFacturados,
                flagCombinacion = p.FlagCombinacion,
                flagCargo = p.FlagCargo,
                sPrecio,
                finicio = p.FechaInicio,
                ffinal = p.FechaFin
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    // ── KDS ────────────────────────────────────────────────────────────────

    /// <inheritdoc />
    public async Task<IReadOnlyList<TiempoKdsPedidoRow>> ObtenerTiempoKdsPedidoAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte KDS TiempoPedido: {FechaInicio} – {FechaFin}", fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<TiempoKdsPedidoRow>(
            conn,
            "USP_KDS_ResporteTiempoPedido",
            new { fInicio = fechaInicio, fFinal = fechaFin },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<TiempoKdsProductoRow>> ObtenerTiempoKdsProductoAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        string grupo,
        string subGrupo,
        string producto,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte KDS TiempoProducto: {FechaInicio} – {FechaFin}", fechaInicio, fechaFin);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<TiempoKdsProductoRow>(
            conn,
            "USP_KDS_ResporteTiempoProducto",
            new
            {
                fInicio = fechaInicio,
                fFinal = fechaFin,
                tGrupo = grupo,
                tSubGrupo = subGrupo,
                tProducto = producto
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    // ── Anulación / Control de Transacciones ──────────────────────────────

    /// <inheritdoc />
    public async Task<IReadOnlyList<AnulacionRow>> ObtenerAnulacionAsync(
        AnulacionParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation("Reporte Anulacion: {FechaInicio} – {FechaFin} Turno={Turno}",
            p.FechaInicio, p.FechaFin, string.IsNullOrEmpty(p.Turno) ? "(todos)" : p.Turno);
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<AnulacionRow>(
            conn,
            "spRep_Anulacion",
            new
            {
                lFranjaHoraria = p.FranjaHoraria,
                tTurno = p.Turno,
                fInicio = p.FechaInicio,
                fFinal = p.FechaFin,
                lFlag1 = p.FlagFacturados,
                lFlag2 = p.FlagAnulados,
                lFlag3 = p.FlagTransferidos,
                sCriterio = p.Criterio,
                ExportaExcel = false
            },
            cancellationToken: ct);
        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<LiquidacionTicketRow>> ObtenerLiquidacionTicketAsync(
        LiquidacionTicketParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation(
            "Reporte LiquidacionTicket: TodosLosTurnos={TodosLosTurnos} Turno={Turno} Usuario={Usuario}",
            p.TodosLosTurnos,
            string.IsNullOrWhiteSpace(p.Turno) ? "(todos)" : p.Turno,
            string.IsNullOrWhiteSpace(p.Usuario) ? "(todos)" : p.Usuario);

        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var result = await _spExecutor.QueryAsync<LiquidacionTicketRow>(
            conn,
            "spRep_LiquidacionSuma",
            new
            {
                flagTurno = p.TodosLosTurnos,
                flagDiaContable = p.DiaContable,
                sTurno = p.Turno,
                sUsuario = p.Usuario,
                finicio = p.FechaInicio,
                ffinal = p.FechaFin,
                sSectorVenta = p.SectorVenta
            },
            cancellationToken: ct);

        return result.ToList().AsReadOnly();
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<PaloteoTicketRow>> ObtenerPaloteoTicketAsync(
        PaloteoTicketParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation(
            "Reporte PaloteoTicket: Origen={Origen} TodosTurnos={TodosTurnos} Turno={Turno}",
            p.Origen, p.TodosTurnos, string.IsNullOrWhiteSpace(p.Turno) ? "(todos)" : p.Turno);

        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var sqlParams = new DynamicParameters();

        var filtrosPedido = new List<string>();
        var filtrosDocumento = new List<string>();

        if (p.TodosTurnos)
        {
            filtrosPedido.Add("MPEDIDO.fRegistro >= @FechaInicio AND MPEDIDO.fRegistro <= @FechaFin");
            filtrosDocumento.Add("MDOCUMENTO.fRegistro >= @FechaInicio AND MDOCUMENTO.fRegistro <= @FechaFin");
            sqlParams.Add("FechaInicio", p.FechaInicio);
            sqlParams.Add("FechaFin", p.FechaFin);
        }
        else
        {
            filtrosPedido.Add("MPEDIDO.tTurno = @Turno");
            filtrosDocumento.Add("MDOCUMENTO.tTurno = @Turno");
            sqlParams.Add("Turno", p.Turno);
        }

        AddFiltro("MPEDIDO.tSalon", "Salon", p.Salon);
        AddFiltro("vProducto.TipoProducto", "TipoProducto", p.TipoProducto);
        AddFiltro("MPEDIDO.tMozo", "Mozo", p.Mozo);
        AddFiltro("MPEDIDO.tTipoPedido", "TipoPedido", p.TipoPedido);
        AddFiltro("MPEDIDO.CodigoOrigenVenta", "OrigenVenta", p.OrigenVenta);
        AddFiltro("vProducto.Area", "Area", p.Area);
        AddFiltro("vProducto.Grupo", "Grupo", p.Grupo);
        AddFiltro("vProducto.SubGrupo", "SubGrupo", p.SubGrupo);
        AddFiltro("DPEDIDO.tCodigoProducto", "CodigoProducto", p.CodigoProducto);

        sqlParams.Add("Boton2", p.Boton2);
        sqlParams.Add("Boton3", p.Boton3);
        sqlParams.Add("Boton4", p.Boton4);
        sqlParams.Add("Boton5", p.Boton5);

        var wherePedido = string.Join(" AND ", filtrosPedido);
        var whereDocumento = string.Join(" AND ", filtrosDocumento);
        var whereCliente = string.IsNullOrWhiteSpace(p.CodigoCliente)
            ? string.Empty
            : " AND MDOCUMENTO.tCodigoCliente = @CodigoCliente";
        if (!string.IsNullOrWhiteSpace(p.CodigoCliente))
            sqlParams.Add("CodigoCliente", p.CodigoCliente);

        var origenSql = p.Origen switch
        {
            OrigenPaloteoTicket.Produccion => $"""
                SELECT vProducto.Codigo AS tCodProducto, vSalon.tLocal, vSalon.Descripcion AS Salon,
                       vProducto.TipoProducto, vProducto.Grupo, vProducto.SubGrupo, vProducto.Descripcion AS Producto,
                       DPEDIDO.nCantidad AS Cantidad, DPEDIDO.nVenta AS Venta, MPEDIDO.tTipoPedido
                FROM dbo.DPEDIDO
                INNER JOIN dbo.MPEDIDO ON DPEDIDO.tCodigoPedido = MPEDIDO.tCodigoPedido
                LEFT JOIN dbo.vProducto ON DPEDIDO.tCodigoProducto = vProducto.Codigo
                LEFT JOIN dbo.vSalon ON MPEDIDO.tSalon = vSalon.Codigo
                WHERE MPEDIDO.tEstadoPedido <> '03'
                  AND DPEDIDO.tEstadoItem = 'N'
                  AND {wherePedido}
                """,
            OrigenPaloteoTicket.Venta => $"""
                SELECT vProducto.Codigo AS tCodProducto, vSalon.tLocal, vSalon.Descripcion AS Salon,
                       vProducto.TipoProducto, vProducto.Grupo, vProducto.SubGrupo, vProducto.Descripcion AS Producto,
                       DPEDIDO.nCantidad AS Cantidad, DPEDIDO.nVenta AS Venta, MPEDIDO.tTipoPedido
                FROM dbo.DPEDIDO
                INNER JOIN dbo.MPEDIDO ON DPEDIDO.tCodigoPedido = MPEDIDO.tCodigoPedido
                LEFT JOIN dbo.vProducto ON DPEDIDO.tCodigoProducto = vProducto.Codigo
                LEFT JOIN dbo.vSalon ON MPEDIDO.tSalon = vSalon.Codigo
                INNER JOIN dbo.MDOCUMENTO ON DPEDIDO.tDocumento = MDOCUMENTO.tDocumento
                WHERE MPEDIDO.tEstadoPedido <> '03'
                  AND DPEDIDO.tEstadoItem = 'N'
                  AND DPEDIDO.tFacturado IN ('P','F')
                  AND {wherePedido}
                  AND {whereDocumento}
                  {whereCliente}
                """,
            OrigenPaloteoTicket.Cortesia => $"""
                SELECT vProducto.Codigo AS tCodProducto, vSalon.tLocal, vSalon.Descripcion AS Salon,
                       vProducto.TipoProducto, vProducto.Grupo, vProducto.SubGrupo, vProducto.Descripcion AS Producto,
                       DPEDIDO.nCantidad AS Cantidad, DPEDIDO.nVenta AS Venta, MPEDIDO.tTipoPedido
                FROM dbo.DPEDIDO
                INNER JOIN dbo.MPEDIDO ON DPEDIDO.tCodigoPedido = MPEDIDO.tCodigoPedido
                LEFT JOIN dbo.vProducto ON DPEDIDO.tCodigoProducto = vProducto.Codigo
                LEFT JOIN dbo.vSalon ON MPEDIDO.tSalon = vSalon.Codigo
                WHERE MPEDIDO.tEstadoPedido <> '03'
                  AND DPEDIDO.tEstadoItem = 'N'
                  AND DPEDIDO.tFacturado = 'C'
                  AND {wherePedido}
                """,
            OrigenPaloteoTicket.CuentaCorriente => $"""
                SELECT vProducto.Codigo AS tCodProducto, vSalon.tLocal, vSalon.Descripcion AS Salon,
                       vProducto.TipoProducto, vProducto.Grupo, vProducto.SubGrupo, vProducto.Descripcion AS Producto,
                       DPEDIDO.nCantidad AS Cantidad, DPEDIDO.nVenta AS Venta, MPEDIDO.tTipoPedido
                FROM dbo.DPEDIDO
                INNER JOIN dbo.MPEDIDO ON DPEDIDO.tCodigoPedido = MPEDIDO.tCodigoPedido
                LEFT JOIN dbo.vProducto ON DPEDIDO.tCodigoProducto = vProducto.Codigo
                LEFT JOIN dbo.vSalon ON MPEDIDO.tSalon = vSalon.Codigo
                WHERE ISNULL(MPEDIDO.tClienteCtaCte, '') <> ''
                  AND DPEDIDO.tEstadoItem = 'N'
                  AND {wherePedido}
                  AND DPEDIDO.tCodigoPedido NOT IN (
                        SELECT DISTINCT DP2.tCodigoPedido
                        FROM dbo.DPEDIDO DP2
                        INNER JOIN dbo.MDOCUMENTO ON DP2.tDocumento = MDOCUMENTO.tDocumento
                        WHERE ISNULL(DP2.tCodigoPedido, '0') <> '0'
                          AND {whereDocumento}
                  )
                """,
            OrigenPaloteoTicket.Combinacion => $"""
                SELECT vProducto.Codigo AS tCodProducto, vSalon.tLocal, vSalon.Descripcion AS Salon,
                       vProducto.TipoProducto, vProducto.Grupo, vProducto.SubGrupo, vProducto.Descripcion AS Producto,
                       CPEDIDO.nCantidad AS Cantidad, vProducto.nPrecioVenta AS Venta, MPEDIDO.tTipoPedido
                FROM dbo.CPEDIDO
                INNER JOIN dbo.DPEDIDO ON CPEDIDO.tCodigoPedido = DPEDIDO.tCodigoPedido AND CPEDIDO.tItem = DPEDIDO.tItem
                INNER JOIN dbo.MPEDIDO ON CPEDIDO.tCodigoPedido = MPEDIDO.tCodigoPedido
                LEFT JOIN dbo.vSalon ON MPEDIDO.tSalon = vSalon.Codigo
                LEFT JOIN dbo.vProducto ON CPEDIDO.tProductocombo = vProducto.Codigo
                WHERE MPEDIDO.tEstadoPedido <> '03'
                  AND {wherePedido}
                """,
            OrigenPaloteoTicket.Cargos => $"""
                SELECT vProducto.Codigo AS tCodProducto, vSalon.tLocal, vSalon.Descripcion AS Salon,
                       vProducto.TipoProducto, vProducto.Grupo, vProducto.SubGrupo, vProducto.Descripcion AS Producto,
                       DPEDIDO.nCantidad AS Cantidad, DPEDIDO.nVenta AS Venta, MPEDIDO.tTipoPedido
                FROM dbo.DPEDIDO
                INNER JOIN dbo.MPEDIDO ON DPEDIDO.tCodigoPedido = MPEDIDO.tCodigoPedido
                LEFT JOIN dbo.vProducto ON DPEDIDO.tCodigoProducto = vProducto.Codigo
                LEFT JOIN dbo.vSalon ON MPEDIDO.tSalon = vSalon.Codigo
                WHERE MPEDIDO.tEstadoPedido = '05'
                  AND DPEDIDO.tEstadoItem = 'N'
                  AND {wherePedido}
                """,
            _ => $"""
                SELECT vProducto.Codigo AS tCodProducto, vSalon.tLocal, vSalon.Descripcion AS Salon,
                       vProducto.TipoProducto, vProducto.Grupo, vProducto.SubGrupo, vProducto.Descripcion AS Producto,
                       DPEDIDO.nCantidad AS Cantidad, DPEDIDO.nVenta AS Venta, MPEDIDO.tTipoPedido
                FROM dbo.DPEDIDO
                INNER JOIN dbo.MPEDIDO ON DPEDIDO.tCodigoPedido = MPEDIDO.tCodigoPedido
                LEFT JOIN dbo.vProducto ON DPEDIDO.tCodigoProducto = vProducto.Codigo
                LEFT JOIN dbo.vSalon ON MPEDIDO.tSalon = vSalon.Codigo
                INNER JOIN dbo.MDOCUMENTO ON DPEDIDO.tDocumento = MDOCUMENTO.tDocumento
                WHERE MPEDIDO.tEstadoPedido <> '03'
                  AND DPEDIDO.tEstadoItem = 'N'
                  AND DPEDIDO.tFacturado IN ('P','F')
                  AND {wherePedido}
                  {whereCliente}
                """
        };

        var orderBy = p.OrdenarPorCodigoProducto ? "tCodProducto" : "Producto";
        var sql = $"""
            WITH Base AS (
                {origenSql}
            )
            SELECT
                ISNULL(Base.tCodProducto, '') AS TCodProducto,
                MAX(ISNULL(Base.tLocal, '')) AS TLocal,
                MAX(COALESCE(vLocal.Descripcion, Base.tLocal, '')) AS Local,
                MAX(
                    CASE Base.tTipoPedido
                        WHEN '02' THEN COALESCE(NULLIF(@Boton2, ''), ISNULL(Base.Salon, 'Sin Salon'))
                        WHEN '03' THEN COALESCE(NULLIF(@Boton3, ''), ISNULL(Base.Salon, 'Sin Salon'))
                        WHEN '04' THEN COALESCE(NULLIF(@Boton4, ''), ISNULL(Base.Salon, 'Sin Salon'))
                        WHEN '05' THEN COALESCE(NULLIF(@Boton5, ''), ISNULL(Base.Salon, 'Sin Salon'))
                        ELSE ISNULL(Base.Salon, 'Sin Salon')
                    END
                ) AS Salon,
                MAX(ISNULL(Base.TipoProducto, '')) AS TipoProducto,
                ISNULL(Base.Grupo, '') AS Grupo,
                ISNULL(Base.SubGrupo, '') AS SubGrupo,
                ISNULL(Base.Producto, '') AS Producto,
                SUM(ISNULL(Base.Cantidad, 0)) AS Cantidad,
                SUM(ISNULL(Base.Venta, 0)) AS Venta
            FROM Base
            LEFT JOIN dbo.vLocal ON Base.tLocal = vLocal.Codigo
            GROUP BY Base.tCodProducto, Base.Grupo, Base.SubGrupo, Base.Producto
            ORDER BY {orderBy};
            """;

        var result = await conn.QueryAsync<PaloteoTicketRow>(new CommandDefinition(sql, sqlParams, cancellationToken: ct));
        return result.ToList().AsReadOnly();

        void AddFiltro(string campoSql, string parametro, string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return;
            filtrosPedido.Add($"{campoSql} = @{parametro}");
            sqlParams.Add(parametro, valor);
        }
    }

    // ── Delivery Ticket — Cierre de Cajeros Delivery ─────────────────────────

    /// <inheritdoc />
    public async Task<IReadOnlyList<DeliveryTicketRow>> ObtenerDeliveryTicketAsync(
        DeliveryTicketParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation(
            "Reporte DeliveryTicket: TodosTurnos={TodosTurnos} Turno={Turno} TodasCajas={TodasCajas} TodosMotorizados={TodosMotorizados}",
            p.TodosTurnos,
            string.IsNullOrWhiteSpace(p.Turno) ? "(todos)" : p.Turno,
            p.TodasLasCajas,
            p.TodosLosMotorizados);

        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var sqlParams = new DynamicParameters();

        // Criterio de tiempo (turno o rango de fechas)
        string criterioPrincipal;
        if (!p.TodosTurnos)
        {
            criterioPrincipal = "MDOCUMENTO.tTurno = @Turno";
            sqlParams.Add("Turno", p.Turno);
        }
        else
        {
            criterioPrincipal = "MDOCUMENTO.fRegistro >= @FechaInicio AND MDOCUMENTO.fRegistro <= @FechaFin";
            sqlParams.Add("FechaInicio", p.FechaInicio);
            sqlParams.Add("FechaFin", p.FechaFin);
        }

        // Filtro caja opcional
        var filtrosExtra = new List<string>();
        if (!p.TodasLasCajas && !string.IsNullOrWhiteSpace(p.Caja))
        {
            filtrosExtra.Add("MDOCUMENTO.tCaja = @Caja");
            sqlParams.Add("Caja", p.Caja);
        }

        // Filtro motorizado opcional
        if (!p.TodosLosMotorizados && !string.IsNullOrWhiteSpace(p.Motorizado))
        {
            filtrosExtra.Add("MPEDIDO.tMotorizado = @Motorizado");
            sqlParams.Add("Motorizado", p.Motorizado);
        }

        var whereSuffix = filtrosExtra.Count > 0
            ? " AND " + string.Join(" AND ", filtrosExtra)
            : string.Empty;

        // Query idéntica al Legacy (frmRepDeliveryTicket.frm Sub Genera)
        var sql = $"""
            SELECT dbo.MDOCUMENTO.tCaja,
                   dbo.DPREPAGO.tTipoPago,
                   dbo.vTipoPago.Descripcion AS TipoPago,
                   dbo.MPEDIDO.tMotorizado,
                   dbo.vMotorizado.Descripcion AS Motorizado,
                   dbo.DPREPAGO.tDocumento,
                   dbo.MDOCUMENTO.fRegistro,
                   dbo.MDOCUMENTO.nVenta,
                   dbo.MDOCUMENTO.tTurno,
                   dbo.MDOCUMENTO.tUsuario,
                   dbo.DPREPAGO.tMoneda,
                   dbo.vMoneda.tResumido AS Mon,
                   dbo.DPREPAGO.nTipoCambio,
                   dbo.DPREPAGO.nMonto,
                   dbo.DPREPAGO.nVuelto,
                   dbo.TTARJETACREDITO.tDetallado AS Tarjeta,
                   dbo.DPREPAGO.tNumero,
                   dbo.vTipoCancelacion.Descripcion AS OtroTipo
            FROM dbo.MPEDIDO
            LEFT OUTER JOIN dbo.vMotorizado
                ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo
            RIGHT OUTER JOIN
                (SELECT DISTINCT tDocumento, tCodigoPedido FROM DDOCUMENTO) T1
            INNER JOIN dbo.vMoneda
            INNER JOIN dbo.vTipoPago
            INNER JOIN dbo.DPREPAGO
                ON dbo.vTipoPago.Codigo = dbo.DPREPAGO.tTipoPago
                ON dbo.vMoneda.Codigo = dbo.DPREPAGO.tMoneda
            LEFT OUTER JOIN dbo.MDOCUMENTO
                ON dbo.DPREPAGO.tDocumento = dbo.MDOCUMENTO.tDocumento
                ON T1.tDocumento = dbo.DPREPAGO.tDocumento
                ON dbo.MPEDIDO.tCodigoPedido = T1.tCodigoPedido COLLATE Modern_Spanish_CI_AS
            LEFT OUTER JOIN dbo.TTARJETACREDITO
                ON dbo.DPREPAGO.tTarjeta = dbo.TTARJETACREDITO.tCodigoTarjeta
            LEFT OUTER JOIN dbo.vTipoCancelacion
                ON dbo.DPREPAGO.tOtroTipoPago = dbo.vTipoCancelacion.Codigo
            WHERE tTipoPedido = '02'
              AND tEstadoDocumento = '01'
              AND {criterioPrincipal}
              {whereSuffix}
            ORDER BY dbo.MDOCUMENTO.tCaja,
                     dbo.MPEDIDO.tMotorizado,
                     dbo.DPREPAGO.tTipoPago,
                     dbo.DPREPAGO.tMoneda,
                     dbo.MDOCUMENTO.tDocumento
            """;

        var result = await conn.QueryAsync<DeliveryTicketRow>(sql, sqlParams, commandTimeout: 120);
        return result.ToList().AsReadOnly();
    }

    // ── Reservas — Reporte de Reservas ───────────────────────────────────────

    /// <inheritdoc />
    public async Task<IReadOnlyList<ReservaReporteRow>> ObtenerReservasReporteAsync(
        ReservaReporteParametros p,
        CancellationToken ct = default)
    {
        _logger.LogInformation(
            "Reporte Reservas: FechaInicio={FechaInicio} FechaFin={FechaFin} Generado={Generado} Atendido={Atendido} Anulado={Anulado} Orden={Orden}",
            p.FechaHoraInicio, p.FechaHoraFin,
            p.EstadoGenerado, p.EstadoAtendido, p.EstadoAnulado, p.Orden);

        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var sqlParams = new DynamicParameters();
        sqlParams.Add("fInicio", p.FechaHoraInicio);
        sqlParams.Add("fFinal", p.FechaHoraFin);

        // Filtro de estados equivalente al Legacy chkEstado
        var estadosFiltro = new List<string>();
        if (p.EstadoGenerado) estadosFiltro.Add("'01'");
        if (p.EstadoAtendido) estadosFiltro.Add("'02'");
        if (p.EstadoAnulado)  estadosFiltro.Add("'03'");

        var whereEstado = estadosFiltro.Count > 0
            ? $" AND TRESERVA.tEstadoReserva IN({string.Join(", ", estadosFiltro)})"
            : string.Empty;

        // Columna de ordenamiento — whitelist para evitar SQL injection
        var orderBy = p.Orden switch
        {
            OrdenReserva.Nombre    => "TRESERVA.tNombre",
            OrdenReserva.Telefono  => "TRESERVA.tTelefono",
            OrdenReserva.Fecha     => "FFecha",
            OrdenReserva.Pax       => "TRESERVA.nPax",
            OrdenReserva.Estado    => "TRESERVA.tEstadoReserva",
            _                      => "TRESERVA.tReserva"
        };

        // Query idéntica al Legacy (frmRepReservas.frm Sub ObtenerReservas)
        var sql = $"""
            SELECT TRESERVA.tReserva,
                   TRESERVA.fFecha + TRESERVA.fHora AS FFecha,
                   TRESERVA.tApellido,
                   TRESERVA.tApellido + ' ' + TRESERVA.tNombre AS Cliente,
                   TRESERVA.tNombre,
                   TRESERVA.tTelefono,
                   TRESERVA.nPax,
                   TRESERVA.tEstadoReserva,
                   TRESERVA.tObservacion,
                   TRESERVA.fRegistro,
                   UPPER(vEstadoReserva.Descripcion) AS EstadoReserva
            FROM TRESERVA
            LEFT JOIN vEstadoReserva ON TRESERVA.tEstadoReserva = vEstadoReserva.Codigo
            WHERE (TRESERVA.fFecha + TRESERVA.fHora BETWEEN @fInicio AND @fFinal)
            {whereEstado}
            ORDER BY {orderBy} ASC
            """;

        var result = await conn.QueryAsync<ReservaReporteRow>(new CommandDefinition(sql, sqlParams, cancellationToken: ct));
        return result.ToList().AsReadOnly();
    }
}
