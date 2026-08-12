namespace Inforest.Application.Reportes;

/// <summary>
/// Repositorio de reportes — expone los SPs <c>spRep_*</c> y <c>USP_KDS_Resporte*</c> del Legacy
/// a la capa Application sin lógica de presentación.
/// <para>
/// Todos los SPs se mantienen en SQL Server sin modificación (ADR-002, ADR-007).
/// </para>
/// Reglas: BR-REP-001 … BR-REP-012
/// </summary>
public interface IReporteRepository
{
    // ── Grupo A — Core Operativo ──────────────────────────────────────────────

    /// <summary>
    /// Ejecuta <c>spRep_Comanda</c>.
    /// Legacy: <c>frmRepComanda.frm</c>
    /// Regla: BR-REP-001
    /// </summary>
    Task<IReadOnlyList<ComandaRow>> ObtenerComandaAsync(
        bool flagTipo,
        string orden,
        DateTime fechaInicio,
        DateTime fechaFin,
        string criterio,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_Propina</c>.
    /// Legacy: <c>frmRepPropina.frm</c>
    /// Regla: BR-REP-002
    /// </summary>
    Task<IReadOnlyList<PropinaRow>> ObtenerPropinaAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        string condicion,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_PrincipalCliente</c> (modo detalle).
    /// Legacy: <c>frmRepPrincipal.frm</c>
    /// Regla: BR-REP-003
    /// </summary>
    Task<IReadOnlyList<PrincipalClienteDetalleRow>> ObtenerPrincipalClienteDetalleAsync(
        double montoMinimo,
        string codigoCliente,
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_PrincipalCliente</c> (modo resumen).
    /// Legacy: <c>frmRepPrincipal.frm</c>
    /// Regla: BR-REP-003
    /// </summary>
    Task<IReadOnlyList<PrincipalClienteResumenRow>> ObtenerPrincipalClienteResumenAsync(
        double montoMinimo,
        string codigoCliente,
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_PaloteoComparativo</c>.
    /// Legacy: <c>frmRepPaloteoComparativo.frm</c>
    /// Regla: BR-REP-005
    /// </summary>
    Task<IReadOnlyList<PaloteoComparativoRow>> ObtenerPaloteoComparativoAsync(
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
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_PaloteoSubProd</c>.
    /// Legacy: <c>frmRepPaloteoSubProd.frm</c>
    /// Regla: BR-REP-006
    /// </summary>
    Task<IReadOnlyList<PaloteoSubProdRow>> ObtenerPaloteoSubProdAsync(
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
        CancellationToken ct = default);

    // ── Grupo B — Integrados Multi-Local ─────────────────────────────────────

    /// <summary>
    /// Ejecuta <c>spRep_CtaCteIntegrado</c>.
    /// Legacy: <c>frmRepCtaCteIntegrado.frm</c>
    /// Regla: BR-REP-004
    /// </summary>
    Task<IReadOnlyList<CtaCteIntegradoRow>> ObtenerCtaCteIntegradoAsync(
        string flagTipo,
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_PaloteoVentaIntegrado</c>.
    /// Legacy: <c>frmRepPaloteoVentaIntegrado.frm</c>
    /// Regla: BR-REP-007
    /// </summary>
    Task<IReadOnlyList<PaloteoVentaIntegradoRow>> ObtenerPaloteoVentaIntegradoAsync(
        PaloteoVentaIntegradoParametros parametros,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_RankingIntegrado</c>.
    /// Legacy: <c>frmRepRankingIntegrado.frm</c>
    /// Regla: BR-REP-008
    /// </summary>
    Task<IReadOnlyList<RankingIntegradoRow>> ObtenerRankingIntegradoAsync(
        RankingIntegradoParametros parametros,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_VentaMensualIntegrado</c>.
    /// Legacy: <c>frmRepVentaMensualIntegrado.frm</c>, <c>frmRepDiarioVentaIntegrado.frm</c>
    /// Regla: BR-REP-009
    /// </summary>
    Task<IReadOnlyList<VentaMensualIntegradoRow>> ObtenerVentaMensualIntegradoAsync(
        VentaMensualIntegradoParametros parametros,
        CancellationToken ct = default);

    // ── Grupo C — Delivery / Motorizado ──────────────────────────────────────

    /// <summary>
    /// Ejecuta <c>spRep_AnaliticoMotorizadoIntegrado</c>.
    /// Legacy: <c>frmRepAnaliticoMotorizadoIntegrado.frm</c>
    /// Regla: BR-REP-010
    /// </summary>
    Task<IReadOnlyList<AnaliticoMotorizadoIntegradoRow>> ObtenerAnaliticoMotorizadoIntegradoAsync(
        AnaliticoMotorizadoIntegradoParametros parametros,
        CancellationToken ct = default);

    // ── KDS ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Ejecuta <c>USP_KDS_ResporteTiempoPedido</c>.
    /// Regla: BR-REP-011
    /// </summary>
    Task<IReadOnlyList<TiempoKdsPedidoRow>> ObtenerTiempoKdsPedidoAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>USP_KDS_ResporteTiempoProducto</c>.
    /// Regla: BR-REP-012
    /// </summary>
    Task<IReadOnlyList<TiempoKdsProductoRow>> ObtenerTiempoKdsProductoAsync(
        DateTime fechaInicio,
        DateTime fechaFin,
        string grupo,
        string subGrupo,
        string producto,
        CancellationToken ct = default);
}

// ── Clases de parámetros para SPs complejos ───────────────────────────────────

/// <summary>Parámetros para <c>spRep_PaloteoVentaIntegrado</c>. Regla: BR-REP-007</summary>
public sealed class PaloteoVentaIntegradoParametros
{
    public bool FlagPropiedades { get; init; }
    public bool FlagProduccion { get; init; }
    public bool FlagVenta { get; init; }
    public bool FlagCortesia { get; init; }
    public bool FlagCuentaCte { get; init; }
    public bool FlagPedidosFacturados { get; init; }
    public bool FlagCombinacion { get; init; }
    public bool FlagCargo { get; init; }
    /// <summary>true=detallado, false=resumido</summary>
    public bool FlagOpcion { get; init; }
    public ExpresionPrecio TipoPrecio { get; init; } = ExpresionPrecio.Venta;
    public string TipoProducto { get; init; } = string.Empty;
    public string TipoPedido { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Orden { get; init; } = string.Empty;
    public string Boton2 { get; init; } = string.Empty;
    public string Boton3 { get; init; } = string.Empty;
    public string Boton4 { get; init; } = string.Empty;
    public string Boton5 { get; init; } = string.Empty;
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }
}

/// <summary>Parámetros para <c>spRep_RankingIntegrado</c>. Regla: BR-REP-008</summary>
public sealed class RankingIntegradoParametros
{
    public bool FlagProduccion { get; init; }
    public bool FlagVenta { get; init; }
    public bool FlagCortesia { get; init; }
    public bool FlagCuentaCte { get; init; }
    public bool FlagPedidosFacturados { get; init; }
    public bool FlagCombinacion { get; init; }
    public bool FlagCargo { get; init; }
    /// <summary>true=detallado, false=resumido</summary>
    public bool FlagOpcion { get; init; }
    public ExpresionPrecio TipoPrecio { get; init; } = ExpresionPrecio.Venta;
    public string TipoProducto { get; init; } = string.Empty;
    public string TipoPedido { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Orden { get; init; } = string.Empty;
    public string Boton2 { get; init; } = string.Empty;
    public string Boton3 { get; init; } = string.Empty;
    public string Boton4 { get; init; } = string.Empty;
    public string Boton5 { get; init; } = string.Empty;
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }
}

/// <summary>Parámetros para <c>spRep_VentaMensualIntegrado</c>. Regla: BR-REP-009</summary>
public sealed class VentaMensualIntegradoParametros
{
    public ExpresionPrecio TipoPrecio { get; init; } = ExpresionPrecio.Venta;
    public string Ano { get; init; } = string.Empty;
    public string Mes { get; init; } = string.Empty;
    /// <summary>Condición fecha adicional (vacío = sin restricción extra)</summary>
    public string CondicionFecha { get; init; } = string.Empty;
    /// <summary>Hora de corte para día contable (ej. 6 = 6:00 AM)</summary>
    public double HoraCorte { get; init; }
    /// <summary>'D'=diario, 'M'=mensual</summary>
    public string Tipo { get; init; } = "D";
    public double MetaMensual { get; init; }
    public double DiasDelMes { get; init; }
}

/// <summary>Parámetros para <c>spRep_AnaliticoMotorizadoIntegrado</c>. Regla: BR-REP-010</summary>
public sealed class AnaliticoMotorizadoIntegradoParametros
{
    public bool FlagProduccion { get; init; }
    public bool FlagVenta { get; init; }
    public bool FlagCortesia { get; init; }
    public bool FlagCuentaCte { get; init; }
    public bool FlagPedidosFacturados { get; init; }
    public bool FlagCombinacion { get; init; }
    public bool FlagCargo { get; init; }
    public ExpresionPrecio TipoPrecio { get; init; } = ExpresionPrecio.Venta;
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }
}
