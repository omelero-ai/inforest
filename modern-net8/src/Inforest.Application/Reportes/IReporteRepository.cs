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
    /// Ejecuta <c>spRep_CtaCteN</c>.
    /// Legacy: <c>frmRepCtaCte.frm</c>
    /// Regla: BR-REP-013
    /// </summary>
    Task<IReadOnlyList<CtaCteOperativaRow>> ObtenerCtaCteOperativaAsync(
        CtaCteOperativaParametros parametros,
        CancellationToken ct = default);

    /// <summary>
    /// Obtiene el catálogo activo de tipos de cuenta corriente.
    /// Legacy: <c>vTipoCtaCte</c>, <c>frmRepCtaCte.frm</c>
    /// </summary>
    Task<IReadOnlyList<ReporteFiltroOpcion>> ObtenerTiposCtaCteAsync(CancellationToken ct = default);

    /// <summary>
    /// Obtiene el catálogo activo de subtipos de cuenta corriente.
    /// Legacy: <c>vSubTipoCtaCte</c>, <c>frmRepCtaCte.frm</c>
    /// </summary>
    Task<IReadOnlyList<ReporteFiltroOpcion>> ObtenerSubTiposCtaCteAsync(
        string tipoCtaCte,
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

    /// <summary>Ejecuta <c>spRep_Anulacion</c>. Regla: BR-REP-014</summary>
    Task<IReadOnlyList<AnulacionRow>> ObtenerAnulacionAsync(
        AnulacionParametros parametros,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta <c>spRep_LiquidacionSuma</c>.
    /// Legacy: <c>frmRepLiquidacionTicket.frm</c>
    /// Regla: BR-REP-015
    /// </summary>
    Task<IReadOnlyList<LiquidacionTicketRow>> ObtenerLiquidacionTicketAsync(
        LiquidacionTicketParametros parametros,
        CancellationToken ct = default);

    /// <summary>
    /// Ejecuta el query de <c>frmRepPaloteoTicket.frm</c>.
    /// Legacy: paloteo por ticketera con filtros opcionales y origen configurable.
    /// Regla: BR-REP-016
    /// </summary>
    Task<IReadOnlyList<PaloteoTicketRow>> ObtenerPaloteoTicketAsync(
        PaloteoTicketParametros parametros,
        CancellationToken ct = default);

    /// <summary>
    /// Obtiene filas del reporte "Cierre de Cajeros Delivery".
    /// Legacy: <c>frmRepDeliveryTicket.frm</c> — query dinámica sobre MDOCUMENTO/DPREPAGO/MPEDIDO.
    /// Regla: BR-REP-017
    /// </summary>
    Task<IReadOnlyList<DeliveryTicketRow>> ObtenerDeliveryTicketAsync(
        DeliveryTicketParametros parametros,
        CancellationToken ct = default);

    /// <summary>
    /// Obtiene filas del reporte de reservas.
    /// Legacy: <c>frmRepReservas.frm</c> — query dinámica sobre TRESERVA + vEstadoReserva.
    /// Regla: BR-REP-018
    /// </summary>
    Task<IReadOnlyList<ReservaReporteRow>> ObtenerReservasReporteAsync(
        ReservaReporteParametros parametros,
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

/// <summary>Parámetros para <c>spRep_CtaCteN</c>. Regla: BR-REP-013</summary>
public sealed class CtaCteOperativaParametros
{
    public bool FlagDetalle { get; init; }
    public bool FlagResumido { get; init; }
    public bool FlagConsolidado { get; init; }
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string Cliente { get; init; } = string.Empty;
    public string TipoCtaCte { get; init; } = string.Empty;
    public string SubTipoCtaCte { get; init; } = string.Empty;
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

/// <summary>
/// Parámetros para <c>spRep_Anulacion</c>. Regla: BR-REP-014
/// Legacy: <c>frmRepAnulado.frm</c> — "Control de Transacciones"
/// </summary>
public sealed class AnulacionParametros
{
    /// <summary>true = filtrar por franja horaria sin importar el día</summary>
    public bool FranjaHoraria { get; init; }
    /// <summary>Código de turno; string.Empty = todos los turnos</summary>
    public string Turno { get; init; } = string.Empty;
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }
    /// <summary>Incluir ítems facturados (tEstadoItem='N')</summary>
    public bool FlagFacturados { get; init; } = true;
    /// <summary>Incluir pedidos anulados (tEstadoPedido='03')</summary>
    public bool FlagAnulados { get; init; } = true;
    /// <summary>Incluir ítems/pedidos transferidos</summary>
    public bool FlagTransferidos { get; init; } = true;
    /// <summary>Criterio SQL adicional construido por la UI (salon, usuario, motivo, estado impresión)</summary>
    public string Criterio { get; init; } = string.Empty;
}

/// <summary>
/// Parámetros para <c>spRep_LiquidacionSuma</c>. Regla: BR-REP-015
/// Legacy: <c>frmRepLiquidacionTicket.frm</c> — "Liquidación de Cajero por Ticketera"
/// </summary>
public sealed class LiquidacionTicketParametros
{
    /// <summary>true = consultar todos los turnos por rango de fechas; false = un turno específico</summary>
    public bool TodosLosTurnos { get; init; }
    /// <summary>true = filtrar por día contable; false = usar <c>MDOCUMENTO.fRegistro</c></summary>
    public bool DiaContable { get; init; }
    /// <summary>Código de turno cuando <see cref="TodosLosTurnos"/> es false.</summary>
    public string Turno { get; init; } = string.Empty;
    /// <summary>Usuario opcional; vacío = todos.</summary>
    public string Usuario { get; init; } = string.Empty;
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }
    /// <summary>Sector de venta opcional; vacío = todos.</summary>
    public string SectorVenta { get; init; } = string.Empty;
}

/// <summary>
/// Origen de datos para paloteo por ticketera (<c>frmRepPaloteoTicket.frm</c>).
/// </summary>
public enum OrigenPaloteoTicket
{
    Produccion = 0,
    Venta = 1,
    Cortesia = 2,
    CuentaCorriente = 3,
    Combinacion = 4,
    Cargos = 5,
    PedidosFacturados = 6
}

/// <summary>
/// Parámetros para consulta de paloteo por ticketera. Regla: BR-REP-016
/// Legacy: <c>frmRepPaloteoTicket.frm</c>
/// </summary>
public sealed class PaloteoTicketParametros
{
    /// <summary>true = filtra por rango de fechas; false = turno específico.</summary>
    public bool TodosTurnos { get; init; } = true;
    public string Turno { get; init; } = string.Empty;
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }

    public string Salon { get; init; } = string.Empty;
    public string TipoProducto { get; init; } = string.Empty;
    public string Mozo { get; init; } = string.Empty;
    public string TipoPedido { get; init; } = string.Empty;
    public string OrigenVenta { get; init; } = string.Empty;
    public string Area { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string CodigoProducto { get; init; } = string.Empty;
    public string CodigoCliente { get; init; } = string.Empty;

    public OrigenPaloteoTicket Origen { get; init; } = OrigenPaloteoTicket.Produccion;
    public bool OrdenarPorCodigoProducto { get; init; }
    public bool MostrarTotalPorProducto { get; init; }

    public string Boton2 { get; init; } = string.Empty;
    public string Boton3 { get; init; } = string.Empty;
    public string Boton4 { get; init; } = string.Empty;
    public string Boton5 { get; init; } = string.Empty;
}
