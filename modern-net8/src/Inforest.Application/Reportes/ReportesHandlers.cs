using Inforest.Application.Interfaces;
using Inforest.Application.Reportes;

namespace Inforest.Application.Reportes;

// ============================================================
// Queries y Handlers para reportes — Etapa 10
// Grupo A (Core) + Grupo B (Integrados) + Grupo C (Delivery/KDS)
// Legacy: frmRep*.frm, 5. SP.sql — spRep_*, USP_KDS_Resporte*
// Reglas: BR-REP-001 … BR-REP-012
// ============================================================

// ── BR-REP-001 — Comanda ─────────────────────────────────────────────────────

/// <summary>Query para reporte de comanda. Legacy: <c>frmRepComanda.frm</c>, <c>spRep_Comanda</c></summary>
public sealed record ObtenerReporteComandaQuery(
    bool FlagTipo,
    string Orden,
    DateTime FechaInicio,
    DateTime FechaFin,
    string Criterio = "");

/// <summary>Handler para <see cref="ObtenerReporteComandaQuery"/>.</summary>
public sealed class ObtenerReporteComandaHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteComandaHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<ComandaRow>> HandleAsync(
        ObtenerReporteComandaQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerComandaAsync(q.FlagTipo, q.Orden, q.FechaInicio, q.FechaFin, q.Criterio, ct);
        return new ReporteResultado<ComandaRow>
        {
            Filas = filas,
            TituloReporte = "Reporte de Comanda",
            NombrePlantilla = q.FlagTipo ? "RepComandaDetallado.frx" : "RepComanda.frx"
        };
    }
}

// ── BR-REP-002 — Propina ─────────────────────────────────────────────────────

/// <summary>
/// Query para reporte de propina.
/// Legacy: <c>frmRepPropina.frm</c>, <c>spRep_Propina</c>
/// Regla: BR-REP-002
/// </summary>
public sealed record ObtenerReportePropinaQuery(
    DateTime FechaInicio,
    DateTime FechaFin,
    string Condicion = "",
    /// <summary>
    /// true = Detallado (dsrPropinaD); false = Resumido (dsrPropinaR).
    /// Legacy: optOpcion(0)=Detallado, optOpcion(1)=Resumido
    /// </summary>
    bool EsDetallado = true);

/// <summary>Handler para <see cref="ObtenerReportePropinaQuery"/>.</summary>
public sealed class ObtenerReportePropinaHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReportePropinaHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<PropinaRow>> HandleAsync(
        ObtenerReportePropinaQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerPropinaAsync(q.FechaInicio, q.FechaFin, q.Condicion, ct);
        return new ReporteResultado<PropinaRow>
        {
            Filas = filas,
            TituloReporte = "Reporte de Propinas",
            NombrePlantilla = q.EsDetallado ? "RepPropina.frx" : "RepPropinaResumido.frx"
        };
    }
}

// ── BR-REP-003 — Principal Cliente ───────────────────────────────────────────

/// <summary>
/// Query para reporte de clientes principales.
/// Legacy: <c>frmRepPrincipal.frm</c>, <c>spRep_PrincipalCliente</c>
/// </summary>
public sealed record ObtenerReportePrincipalClienteQuery(
    bool ModoDetalle,
    double MontoMinimo,
    string CodigoCliente,
    DateTime FechaInicio,
    DateTime FechaFin);

/// <summary>Handler para <see cref="ObtenerReportePrincipalClienteQuery"/>.</summary>
public sealed class ObtenerReportePrincipalClienteHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReportePrincipalClienteHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<object>> HandleAsync(
        ObtenerReportePrincipalClienteQuery q,
        CancellationToken ct = default)
    {
        if (q.ModoDetalle)
        {
            var filas = await _repo.ObtenerPrincipalClienteDetalleAsync(
                q.MontoMinimo, q.CodigoCliente, q.FechaInicio, q.FechaFin, ct);
            return new ReporteResultado<object>
            {
                Filas = filas.Cast<object>().ToList().AsReadOnly(),
                TituloReporte = "Clientes Principales — Detalle",
                NombrePlantilla = "RepPrincipalClienteDetalle.frx"
            };
        }
        else
        {
            var filas = await _repo.ObtenerPrincipalClienteResumenAsync(
                q.MontoMinimo, q.CodigoCliente, q.FechaInicio, q.FechaFin, ct);
            return new ReporteResultado<object>
            {
                Filas = filas.Cast<object>().ToList().AsReadOnly(),
                TituloReporte = "Clientes Principales — Resumen",
                NombrePlantilla = "RepPrincipalClienteResumen.frx"
            };
        }
    }
}

// ── BR-REP-004 — Cuenta Corriente Integrado ──────────────────────────────────

/// <summary>
/// Query para reporte de cuenta corriente integrado.
/// Legacy: <c>frmRepCtaCteIntegrado.frm</c>, <c>spRep_CtaCteIntegrado</c>
/// flagTipo: '1'=consolidado, '2'=detallado, '3'=estado
/// </summary>
public sealed record ObtenerReporteCtaCteIntegradoQuery(
    string FlagTipo,
    DateTime FechaInicio,
    DateTime FechaFin);

/// <summary>Handler para <see cref="ObtenerReporteCtaCteIntegradoQuery"/>.</summary>
public sealed class ObtenerReporteCtaCteIntegradoHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteCtaCteIntegradoHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<CtaCteIntegradoRow>> HandleAsync(
        ObtenerReporteCtaCteIntegradoQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerCtaCteIntegradoAsync(q.FlagTipo, q.FechaInicio, q.FechaFin, ct);
        var plantilla = q.FlagTipo switch
        {
            "1" => "RepCtaCteIntegradoConsolidado.frx",
            "2" => "RepCtaCteIntegradoDetallado.frx",
            _ => "RepCtaCteIntegrado.frx"
        };
        return new ReporteResultado<CtaCteIntegradoRow>
        {
            Filas = filas,
            TituloReporte = "Cuenta Corriente Integrado",
            NombrePlantilla = plantilla
        };
    }
}

// ── BR-REP-005 — Paloteo Comparativo ─────────────────────────────────────────

/// <summary>
/// Query para reporte operativo de cuentas corrientes.
/// Legacy: <c>frmRepCtaCte.frm</c>, <c>spRep_CtaCteN</c>
/// </summary>
public sealed record ObtenerReporteCtaCteOperativaQuery(CtaCteOperativaParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteCtaCteOperativaQuery"/>.</summary>
public sealed class ObtenerReporteCtaCteOperativaHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteCtaCteOperativaHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<CtaCteOperativaRow>> HandleAsync(
        ObtenerReporteCtaCteOperativaQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerCtaCteOperativaAsync(q.Parametros, ct);
        var plantilla = q.Parametros.FlagConsolidado
            ? "RepCtaCteConsolidado.frx"
            : q.Parametros.FlagResumido
                ? "RepCtaCteResumido.frx"
                : "RepCtaCteDetallado.frx";

        return new ReporteResultado<CtaCteOperativaRow>
        {
            Filas = filas,
            TituloReporte = "Estados de Cuentas Corrientes",
            NombrePlantilla = plantilla
        };
    }
}

/// <summary>Handler para catálogo de tipos de cuenta corriente del reporte operativo.</summary>
public sealed class ObtenerTiposCtaCteReporteHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerTiposCtaCteReporteHandler(IReporteRepository repo) => _repo = repo;

    public Task<IReadOnlyList<ReporteFiltroOpcion>> HandleAsync(CancellationToken ct = default)
        => _repo.ObtenerTiposCtaCteAsync(ct);
}

/// <summary>Handler para catálogo de subtipos de cuenta corriente del reporte operativo.</summary>
public sealed class ObtenerSubTiposCtaCteReporteHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerSubTiposCtaCteReporteHandler(IReporteRepository repo) => _repo = repo;

    public Task<IReadOnlyList<ReporteFiltroOpcion>> HandleAsync(string tipoCtaCte, CancellationToken ct = default)
        => _repo.ObtenerSubTiposCtaCteAsync(tipoCtaCte, ct);
}

/// <summary>
/// Query para reporte paloteo comparativo.
/// Legacy: <c>frmRepPaloteoComparativo.frm</c>, <c>spRep_PaloteoComparativo</c>
/// </summary>
public sealed record ObtenerReportePaloteoComparativoQuery(
    bool PorTurno,
    bool PorValor,
    string TipoProducto,
    string AreaProduccion,
    string Grupo,
    string SubGrupo,
    string CodigoProducto,
    string Turno,
    DateTime FechaInicio,
    DateTime FechaFin,
    bool SoloFacturados);

/// <summary>Handler para <see cref="ObtenerReportePaloteoComparativoQuery"/>.</summary>
public sealed class ObtenerReportePaloteoComparativoHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReportePaloteoComparativoHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<PaloteoComparativoRow>> HandleAsync(
        ObtenerReportePaloteoComparativoQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerPaloteoComparativoAsync(
            q.PorTurno, q.PorValor, q.TipoProducto, q.AreaProduccion,
            q.Grupo, q.SubGrupo, q.CodigoProducto, q.Turno,
            q.FechaInicio, q.FechaFin, q.SoloFacturados, ct);
        return new ReporteResultado<PaloteoComparativoRow>
        {
            Filas = filas,
            TituloReporte = "Paloteo Comparativo",
            NombrePlantilla = "RepPaloteoComparativo.frx"
        };
    }
}

// ── BR-REP-006 — Paloteo Sub-Producto ────────────────────────────────────────

/// <summary>
/// Query para reporte paloteo sub-productos.
/// Legacy: <c>frmRepPaloteoSubProd.frm</c>, <c>spRep_PaloteoSubProd</c>
/// </summary>
public sealed record ObtenerReportePaloteoSubProdQuery(
    bool PorTurno,
    string Turno,
    string Local,
    string Salon,
    string TipoProducto,
    string Mozo,
    string TipoPedido,
    string Area,
    string Caja,
    string CodigoProducto,
    string CodigoCliente,
    DateTime FechaInicio,
    DateTime FechaFin,
    string Boton2 = "",
    string Boton3 = "",
    string Boton4 = "",
    string Boton5 = "");

/// <summary>Handler para <see cref="ObtenerReportePaloteoSubProdQuery"/>.</summary>
public sealed class ObtenerReportePaloteoSubProdHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReportePaloteoSubProdHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<PaloteoSubProdRow>> HandleAsync(
        ObtenerReportePaloteoSubProdQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerPaloteoSubProdAsync(
            q.PorTurno, q.Turno, q.Local, q.Salon, q.TipoProducto, q.Mozo,
            q.TipoPedido, q.Area, q.Caja, q.CodigoProducto, q.CodigoCliente,
            q.Boton2, q.Boton3, q.Boton4, q.Boton5,
            q.FechaInicio, q.FechaFin, ct);
        return new ReporteResultado<PaloteoSubProdRow>
        {
            Filas = filas,
            TituloReporte = "Paloteo por Sub-Producto",
            NombrePlantilla = "RepPaloteoSubProd.frx"
        };
    }
}

// ── BR-REP-007 — Paloteo Venta Integrado ─────────────────────────────────────

/// <summary>
/// Query para reporte paloteo de venta integrado.
/// Legacy: <c>frmRepPaloteoVentaIntegrado.frm</c>, <c>spRep_PaloteoVentaIntegrado</c>
/// </summary>
public sealed record ObtenerReportePaloteoVentaIntegradoQuery(PaloteoVentaIntegradoParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReportePaloteoVentaIntegradoQuery"/>.</summary>
public sealed class ObtenerReportePaloteoVentaIntegradoHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReportePaloteoVentaIntegradoHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<PaloteoVentaIntegradoRow>> HandleAsync(
        ObtenerReportePaloteoVentaIntegradoQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerPaloteoVentaIntegradoAsync(q.Parametros, ct);
        var plantilla = q.Parametros.FlagOpcion
            ? "RepPaloteoVentaIntegradoDetallado.frx"
            : "RepPaloteoVentaIntegradoResumido.frx";
        return new ReporteResultado<PaloteoVentaIntegradoRow>
        {
            Filas = filas,
            TituloReporte = "Paloteo de Venta Integrado",
            NombrePlantilla = plantilla
        };
    }
}

// ── BR-REP-008 — Ranking Integrado ───────────────────────────────────────────

/// <summary>
/// Query para reporte ranking integrado.
/// Legacy: <c>frmRepRankingIntegrado.frm</c>, <c>spRep_RankingIntegrado</c>
/// </summary>
public sealed record ObtenerReporteRankingIntegradoQuery(RankingIntegradoParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteRankingIntegradoQuery"/>.</summary>
public sealed class ObtenerReporteRankingIntegradoHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteRankingIntegradoHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<RankingIntegradoRow>> HandleAsync(
        ObtenerReporteRankingIntegradoQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerRankingIntegradoAsync(q.Parametros, ct);
        var plantilla = q.Parametros.FlagOpcion
            ? "RepRankingIntegradoDetallado.frx"
            : "RepRankingIntegradoResumido.frx";
        return new ReporteResultado<RankingIntegradoRow>
        {
            Filas = filas,
            TituloReporte = "Ranking Integrado",
            NombrePlantilla = plantilla
        };
    }
}

// ── BR-REP-009 — Venta Mensual Integrado ─────────────────────────────────────

/// <summary>
/// Query para reporte de venta mensual/diaria integrado.
/// Legacy: <c>frmRepVentaMensualIntegrado.frm</c>, <c>frmRepDiarioVentaIntegrado.frm</c>,
///         <c>spRep_VentaMensualIntegrado</c>
/// </summary>
public sealed record ObtenerReporteVentaMensualIntegradoQuery(VentaMensualIntegradoParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteVentaMensualIntegradoQuery"/>.</summary>
public sealed class ObtenerReporteVentaMensualIntegradoHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteVentaMensualIntegradoHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<VentaMensualIntegradoRow>> HandleAsync(
        ObtenerReporteVentaMensualIntegradoQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerVentaMensualIntegradoAsync(q.Parametros, ct);
        var plantilla = q.Parametros.Tipo == "D"
            ? "RepVentaDiariaIntegrado.frx"
            : "RepVentaMensualIntegrado.frx";
        return new ReporteResultado<VentaMensualIntegradoRow>
        {
            Filas = filas,
            TituloReporte = q.Parametros.Tipo == "D" ? "Venta Diaria Integrada" : "Venta Mensual Integrada",
            NombrePlantilla = plantilla
        };
    }
}

// ── BR-REP-010 — Analítico Motorizado Integrado ───────────────────────────────

/// <summary>
/// Query para reporte analítico de motorizado integrado.
/// Legacy: <c>frmRepAnaliticoMotorizadoIntegrado.frm</c>, <c>spRep_AnaliticoMotorizadoIntegrado</c>
/// </summary>
public sealed record ObtenerReporteAnaliticoMotorizadoIntegradoQuery(
    AnaliticoMotorizadoIntegradoParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteAnaliticoMotorizadoIntegradoQuery"/>.</summary>
public sealed class ObtenerReporteAnaliticoMotorizadoIntegradoHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteAnaliticoMotorizadoIntegradoHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<AnaliticoMotorizadoIntegradoRow>> HandleAsync(
        ObtenerReporteAnaliticoMotorizadoIntegradoQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerAnaliticoMotorizadoIntegradoAsync(q.Parametros, ct);
        return new ReporteResultado<AnaliticoMotorizadoIntegradoRow>
        {
            Filas = filas,
            TituloReporte = "Analítico Motorizado Integrado",
            NombrePlantilla = "RepAnaliticoMotorizadoIntegrado.frx"
        };
    }
}

// ── BR-REP-011 — KDS Tiempo Pedido ───────────────────────────────────────────

/// <summary>
/// Query para reporte de tiempo KDS por pedido.
/// Legacy: <c>USP_KDS_ResporteTiempoPedido</c>
/// </summary>
public sealed record ObtenerReporteTiempoKdsPedidoQuery(DateTime FechaInicio, DateTime FechaFin);

/// <summary>Handler para <see cref="ObtenerReporteTiempoKdsPedidoQuery"/>.</summary>
public sealed class ObtenerReporteTiempoKdsPedidoHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteTiempoKdsPedidoHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<TiempoKdsPedidoRow>> HandleAsync(
        ObtenerReporteTiempoKdsPedidoQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerTiempoKdsPedidoAsync(q.FechaInicio, q.FechaFin, ct);
        return new ReporteResultado<TiempoKdsPedidoRow>
        {
            Filas = filas,
            TituloReporte = "Tiempo KDS por Pedido",
            NombrePlantilla = "RepTiempoKdsPedido.frx"
        };
    }
}

// ── BR-REP-012 — KDS Tiempo Producto ─────────────────────────────────────────

/// <summary>
/// Query para reporte de tiempo KDS por producto.
/// Legacy: <c>USP_KDS_ResporteTiempoProducto</c>
/// </summary>
public sealed record ObtenerReporteTiempoKdsProductoQuery(
    DateTime FechaInicio,
    DateTime FechaFin,
    string Grupo = "",
    string SubGrupo = "",
    string Producto = "");

/// <summary>Handler para <see cref="ObtenerReporteTiempoKdsProductoQuery"/>.</summary>
public sealed class ObtenerReporteTiempoKdsProductoHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteTiempoKdsProductoHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<TiempoKdsProductoRow>> HandleAsync(
        ObtenerReporteTiempoKdsProductoQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerTiempoKdsProductoAsync(
            q.FechaInicio, q.FechaFin, q.Grupo, q.SubGrupo, q.Producto, ct);
        return new ReporteResultado<TiempoKdsProductoRow>
        {
            Filas = filas,
            TituloReporte = "Tiempo KDS por Producto",
            NombrePlantilla = "RepTiempoKdsProducto.frx"
        };
    }
}

// ── BR-REP-014 — Anulación / Control de Transacciones ────────────────────────

/// <summary>
/// Query para reporte de anulaciones y control de transacciones.
/// Legacy: <c>frmRepAnulado.frm</c>, <c>spRep_Anulacion</c>
/// Regla: BR-REP-014
/// </summary>
public sealed record ObtenerReporteAnulacionQuery(AnulacionParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteAnulacionQuery"/>.</summary>
public sealed class ObtenerReporteAnulacionHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteAnulacionHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<AnulacionRow>> HandleAsync(
        ObtenerReporteAnulacionQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerAnulacionAsync(q.Parametros, ct);
        return new ReporteResultado<AnulacionRow>
        {
            Filas = filas,
            TituloReporte = "Control de Transacciones",
            NombrePlantilla = "RepAnulacion.frx"
        };
    }
}

// ── BR-REP-015 — Liquidación de Cajero por Ticketera ──────────────────────────

/// <summary>
/// Query para liquidación de cajero por ticketera.
/// Legacy: <c>frmRepLiquidacionTicket.frm</c>, <c>spRep_LiquidacionSuma</c>
/// Regla: BR-REP-015
/// </summary>
public sealed record ObtenerReporteLiquidacionTicketQuery(LiquidacionTicketParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteLiquidacionTicketQuery"/>.</summary>
public sealed class ObtenerReporteLiquidacionTicketHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteLiquidacionTicketHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<LiquidacionTicketRow>> HandleAsync(
        ObtenerReporteLiquidacionTicketQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerLiquidacionTicketAsync(q.Parametros, ct);
        return new ReporteResultado<LiquidacionTicketRow>
        {
            Filas = filas,
            TituloReporte = "Liquidación de Cajero por Ticketera",
            NombrePlantilla = "RepLiquidacionTicket.frx"
        };
    }
}

// ── BR-REP-016 — Paloteo de Producción por Ticketera ───────────────────────────

/// <summary>
/// Query para paloteo por ticketera.
/// Legacy: <c>frmRepPaloteoTicket.frm</c>
/// Regla: BR-REP-016
/// </summary>
public sealed record ObtenerReportePaloteoTicketQuery(PaloteoTicketParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReportePaloteoTicketQuery"/>.</summary>
public sealed class ObtenerReportePaloteoTicketHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReportePaloteoTicketHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<PaloteoTicketRow>> HandleAsync(
        ObtenerReportePaloteoTicketQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerPaloteoTicketAsync(q.Parametros, ct);
        return new ReporteResultado<PaloteoTicketRow>
        {
            Filas = filas,
            TituloReporte = "Paloteo de Producción por Ticketera",
            NombrePlantilla = "RepPaloteoTicket.frx"
        };
    }
}

// ── DeliveryTicket — Cierre de Cajeros Delivery ───────────────────────────────

/// <summary>
/// Query para obtener el reporte "Cierre de Cajeros Delivery".
/// Legacy: <c>frmRepDeliveryTicket.frm</c>
/// Regla: BR-REP-017
/// </summary>
public sealed record ObtenerReporteDeliveryTicketQuery(DeliveryTicketParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteDeliveryTicketQuery"/>.</summary>
public sealed class ObtenerReporteDeliveryTicketHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteDeliveryTicketHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<DeliveryTicketRow>> HandleAsync(
        ObtenerReporteDeliveryTicketQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerDeliveryTicketAsync(q.Parametros, ct);
        return new ReporteResultado<DeliveryTicketRow>
        {
            Filas = filas,
            TituloReporte = "Cierre de Cajeros Delivery",
            NombrePlantilla = "RepDeliveryTicket.frx"
        };
    }
}

// ── BR-REP-018 — Reporte de Reservas ─────────────────────────────────────────

/// <summary>
/// Query para obtener el reporte de reservas.
/// Legacy: <c>frmRepReservas.frm</c> — query dinámica sobre TRESERVA + vEstadoReserva.
/// Regla: BR-REP-018
/// </summary>
public sealed record ObtenerReporteReservasQuery(ReservaReporteParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteReservasQuery"/>.</summary>
public sealed class ObtenerReporteReservasHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteReservasHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<ReservaReporteRow>> HandleAsync(
        ObtenerReporteReservasQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerReservasReporteAsync(q.Parametros, ct);
        return new ReporteResultado<ReservaReporteRow>
        {
            Filas = filas,
            TituloReporte = "Reservas",
            NombrePlantilla = "RepReservas.frx"
        };
    }
}

// ── BR-REP-019 — Reporte de Entregas ─────────────────────────────────────────

/// <summary>
/// Query para obtener el reporte de entregas.
/// Legacy: <c>frmRepEntrega.frm</c>, <c>spRep_Entregas</c>
/// Regla: BR-REP-019
/// </summary>
public sealed record ObtenerReporteEntregaQuery(EntregaParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteEntregaQuery"/>.</summary>
public sealed class ObtenerReporteEntregaHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteEntregaHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<EntregaRow>> HandleAsync(
        ObtenerReporteEntregaQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerEntregasAsync(q.Parametros, ct);
        return new ReporteResultado<EntregaRow>
        {
            Filas = filas,
            TituloReporte = "Reporte de Entregas",
            NombrePlantilla = q.Parametros.Formato switch
            {
                FormatoReporteEntrega.DetalladoFormato1 => "RepEntregaFormato1.frx",
                FormatoReporteEntrega.DetalladoFormato2 => "RepEntregaFormato2.frx",
                _ => "RepEntregaResumidoProd.frx"
            }
        };
    }
}

// ── BR-REP-020 — Venta Mensual por Fechas ─────────────────────────────────────

/// <summary>
/// Query para reporte comparativo de venta mensual por día.
/// Legacy: <c>frmRepVentaFecha.frm</c>, <c>spRep_VentaFecha</c>
/// Regla: BR-REP-020
/// </summary>
public sealed record ObtenerReporteVentaFechaQuery(VentaFechaParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteVentaFechaQuery"/>.</summary>
public sealed class ObtenerReporteVentaFechaHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteVentaFechaHandler(IReporteRepository repo) => _repo = repo;

    public async Task<ReporteResultado<VentaFechaRow>> HandleAsync(
        ObtenerReporteVentaFechaQuery q,
        CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerVentaFechaAsync(q.Parametros, ct);
        var mesNombre = new System.Globalization.CultureInfo("es-PE").DateTimeFormat.GetMonthName(q.Parametros.Mes);
        var sufijoPrecio = q.Parametros.TipoPrecio == TipoPrecioVentaFecha.Neto
            ? "con Precios Netos"
            : "con Precios de Venta";
        return new ReporteResultado<VentaFechaRow>
        {
            Filas = filas,
            TituloReporte = $"Comparativo del Mes de : {mesNombre} {sufijoPrecio}",
            NombrePlantilla = "RepVentaFecha.frx"
        };
    }
}

// ============================================================
// Liquidación de Cajero — frmRepLiquidacion.frm
// Legacy: spRep_LiquidacionOutPut, spRep_Liquidacion (tipos 1-5), spRep_LiquidacionSuma
// Regla: BR-REP-021
// ============================================================

/// <summary>
/// Query para el reporte Liquidación de Cajero.
/// Legacy: <c>frmRepLiquidacion.frm</c>, cmdOpcion_Click → Sub Genera()
/// Regla: BR-REP-021
/// SQL: <c>spRep_LiquidacionOutPut</c>, <c>spRep_Liquidacion</c>, <c>spRep_LiquidacionSuma</c>
/// </summary>
public sealed record ObtenerReporteLiquidacionQuery(LiquidacionParametros Parametros);

/// <summary>Handler para <see cref="ObtenerReporteLiquidacionQuery"/>.</summary>
public sealed class ObtenerReporteLiquidacionHandler
{
    private readonly IReporteRepository _repo;
    public ObtenerReporteLiquidacionHandler(IReporteRepository repo) => _repo = repo;

    /// <summary>
    /// Ejecuta todos los SPs del Legacy necesarios para el reporte de Liquidación de Cajero
    /// y devuelve el resultado consolidado.
    /// Legacy: Sub Genera() + Sub Genera2() — spRep_LiquidacionOutPut tipos 1-5 + spRep_LiquidacionSuma
    /// Regla: BR-REP-021
    /// </summary>
    public async Task<LiquidacionResultado> HandleAsync(
        ObtenerReporteLiquidacionQuery q,
        CancellationToken ct = default)
    {
        var p = q.Parametros;

        // Ejecutar todos los SPs en paralelo para mejorar rendimiento
        var outputTask = _repo.ObtenerLiquidacionOutputAsync(p, ct);
        var documentosTask = _repo.ObtenerLiquidacionDocumentosAsync(p, ct);
        var sumasTask = _repo.ObtenerLiquidacionSumasGrupoAsync(p, ct);
        var tarjetasTask = _repo.ObtenerLiquidacionTarjetasAsync(p, ct);
        var tiposPedidoTask = _repo.ObtenerLiquidacionTiposPedidoAsync(p, ct);
        var otrosTiposTask = _repo.ObtenerLiquidacionOtrosTiposAsync(p, ct);

        await Task.WhenAll(outputTask, documentosTask, sumasTask, tarjetasTask, tiposPedidoTask, otrosTiposTask);

        // Construir título igual que Legacy (sTitulo en Sub cmdOpcion_Click)
        var titulo = ConstruirTitulo(p);

        return new LiquidacionResultado
        {
            Output = await outputTask,
            Documentos = await documentosTask,
            SumasGrupo = await sumasTask,
            Tarjetas = await tarjetasTask,
            TiposPedido = await tiposPedidoTask,
            OtrosTipos = await otrosTiposTask,
            Titulo = titulo
        };
    }

    /// <summary>
    /// Construye el título descriptivo del reporte equivalente al string sTitulo del Legacy.
    /// Legacy: cmdOpcion_Click — sTitulo = "Turno : " & ... & " Del " & ... & " Al " & ...
    /// </summary>
    private static string ConstruirTitulo(LiquidacionParametros p)
    {
        string rango;
        if (p.ModoFiltro == LiquidacionModoFiltro.PorTurno)
        {
            rango = $"Turno : {p.Turno}";
        }
        else if (p.PorDiaContable)
        {
            rango = $"Por Dia Contable, Todos los Turnos del {p.FechaInicio:dd/MM/yyyy} Al {p.FechaFin:dd/MM/yyyy}";
        }
        else
        {
            rango = $"Turno : Todos los Turnos Del {p.FechaInicio:dd/MM/yyyy} {p.FechaInicio:HH:mm} Hrs Al {p.FechaFin:dd/MM/yyyy} {p.FechaFin:HH:mm} Hrs";
        }

        var usuario = string.IsNullOrEmpty(p.Usuario)
            ? "Usuario : Todos los Usuarios"
            : $"Usuario : {p.Usuario}";

        return $"{rango}\n{usuario}";
    }
}
