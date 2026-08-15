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

/// <summary>Query para reporte de propina. Legacy: <c>frmRepPropina.frm</c>, <c>spRep_Propina</c></summary>
public sealed record ObtenerReportePropinaQuery(
    DateTime FechaInicio,
    DateTime FechaFin,
    string Condicion = "");

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
            NombrePlantilla = "RepPropina.frx"
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
