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
}
