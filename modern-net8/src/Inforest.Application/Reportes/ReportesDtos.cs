namespace Inforest.Application.Reportes;

// ============================================================
// DTOs de resultados de reportes — Etapa 10
// Cada clase mapea exactamente el dataset retornado por el SP Legacy.
// Legacy: 5. SP.sql — spRep_*, USP_KDS_Resporte*
// Reglas: BR-REP-001 … BR-REP-012
// ============================================================

/// <summary>
/// Fila del dataset retornado por <c>spRep_Comanda</c>.
/// Legacy: <c>frmRepComanda.frm</c>
/// Regla: BR-REP-001
/// </summary>
public sealed class ComandaRow
{
    public string TCodigoPedido { get; init; } = string.Empty;
    public string TComanda { get; init; } = string.Empty;
    public string NombreProducto { get; init; } = string.Empty;
    public string Mozo { get; init; } = string.Empty;
    public double NCantidad { get; init; }
    public double PrecioUnitario { get; init; }
    public double PrecioTotal { get; init; }
    public DateTime FFecha { get; init; }
    public string Usuario { get; init; } = string.Empty;
    public string TDocumento { get; init; } = string.Empty;
    public string Estado { get; init; } = string.Empty;
    public string TObservacion { get; init; } = string.Empty;
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_Propina</c>.
/// Legacy: <c>frmRepPropina.frm</c>
/// Regla: BR-REP-002
/// </summary>
public sealed class PropinaRow
{
    public string TMozo { get; init; } = string.Empty;
    public string TMotorizado { get; init; } = string.Empty;
    public string TDetallado { get; init; } = string.Empty;
    public double Propina { get; init; }
    public string TDocumento { get; init; } = string.Empty;
    public string FRegistro { get; init; } = string.Empty;
    public string TCorrelativo { get; init; } = string.Empty;
    public string TTipoPedido { get; init; } = string.Empty;
    public string Trabajador { get; init; } = string.Empty;
    public string TipoPedido { get; init; } = string.Empty;
    public double NFactorRetencion { get; init; }
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_PrincipalCliente</c> (modo detalle).
/// Legacy: <c>frmRepPrincipal.frm</c>
/// Regla: BR-REP-003
/// </summary>
public sealed class PrincipalClienteDetalleRow
{
    public string TCodigoCliente { get; init; } = string.Empty;
    public string TEmpresa { get; init; } = string.Empty;
    public string TDocumento { get; init; } = string.Empty;
    public DateTime FFecha { get; init; }
    public double NNeto { get; init; }
    public double NPrecioImpuesto1 { get; init; }
    public double NPrecioImpuesto2 { get; init; }
    public double NPrecioImpuesto3 { get; init; }
    public double NVenta { get; init; }
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_PrincipalCliente</c> (modo resumen).
/// Legacy: <c>frmRepPrincipal.frm</c>
/// Regla: BR-REP-003
/// </summary>
public sealed class PrincipalClienteResumenRow
{
    public string TCodigoCliente { get; init; } = string.Empty;
    public string TEmpresa { get; init; } = string.Empty;
    public int Cantidad { get; init; }
    public double NNeto { get; init; }
    public double NPrecioImpuesto1 { get; init; }
    public double NPrecioImpuesto2 { get; init; }
    public double NPrecioImpuesto3 { get; init; }
    public double NVenta { get; init; }
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_CtaCteIntegrado</c>.
/// Legacy: <c>frmRepCtaCteIntegrado.frm</c>
/// Regla: BR-REP-004
/// </summary>
public sealed class CtaCteIntegradoRow
{
    public string TCodigoCliente { get; init; } = string.Empty;
    public string TEmpresa { get; init; } = string.Empty;
    public string TDocumento { get; init; } = string.Empty;
    public DateTime FFecha { get; init; }
    public double NNeto { get; init; }
    public double NVenta { get; init; }
    public string TLocal { get; init; } = string.Empty;
    public string TEstado { get; init; } = string.Empty;
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_CtaCteN</c>.
/// Legacy: <c>frmRepCtaCte.frm</c>
/// Regla: BR-REP-013
/// </summary>
public sealed class CtaCteOperativaRow
{
    public string Descripcion { get; init; } = string.Empty;
    public string Identidad { get; init; } = string.Empty;
    public double NConsumo { get; init; }
    public double NLinea { get; init; }
    public double NSaldo { get; init; }
    public string TCodigoPedido { get; init; } = string.Empty;
    public DateTime FFecha { get; init; }
    public string Local { get; init; } = string.Empty;
    public string TEstadoPedido { get; init; } = string.Empty;
    public double NVenta { get; init; }
    public string TClienteCtaCte { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public double NCantidad { get; init; }
    public string TDocumento { get; init; } = string.Empty;
    public string TTipoCtaCte { get; init; } = string.Empty;
    public string TSubTipoCtaCte { get; init; } = string.Empty;
    public double Consumo { get; init; }
    public double Linea { get; init; }
    public double Saldo { get; init; }
    public DateTime Fecha { get; init; }
    public double Suma { get; init; }
}

/// <summary>
/// Opción de catálogo para filtros de reportes.
/// Legacy: <c>vTipoCtaCte</c>, <c>vSubTipoCtaCte</c>
/// </summary>
public sealed class ReporteFiltroOpcion
{
    public string Codigo { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_PaloteoComparativo</c>.
/// Legacy: <c>frmRepPaloteoComparativo.frm</c>
/// Regla: BR-REP-005
/// </summary>
public sealed class PaloteoComparativoRow
{
    public string Codigo { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public double Valor { get; init; }
    public double Produccion { get; init; }
    public double Venta { get; init; }
    public double Cortesia { get; init; }
    public double CtaCte { get; init; }
    public double Canal1 { get; init; }
    public double Canal2 { get; init; }
    public double Canal3 { get; init; }
    public double Canal4 { get; init; }
    public double Canal5 { get; init; }
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_PaloteoSubProd</c>.
/// Legacy: <c>frmRepPaloteoSubProd.frm</c>
/// Regla: BR-REP-006
/// </summary>
public sealed class PaloteoSubProdRow
{
    public string TLocal { get; init; } = string.Empty;
    public string Local { get; init; } = string.Empty;
    public string Salon { get; init; } = string.Empty;
    public string TMesa { get; init; } = string.Empty;
    public string TipoProducto { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public double Cantidad { get; init; }
    public double Venta { get; init; }
    public string Pedido { get; init; } = string.Empty;
    public string Documento { get; init; } = string.Empty;
    public DateTime Fecha { get; init; }
    public string TTipoPedido { get; init; } = string.Empty;
    public string Area { get; init; } = string.Empty;
    public string SubProducto { get; init; } = string.Empty;
    public double CantProd { get; init; }
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_PaloteoVentaIntegrado</c>.
/// Legacy: <c>frmRepPaloteoVentaIntegrado.frm</c>
/// Regla: BR-REP-007
/// </summary>
public sealed class PaloteoVentaIntegradoRow
{
    public string TLocal { get; init; } = string.Empty;
    public string Local { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public double Cantidad { get; init; }
    public double Venta { get; init; }
    public double Produccion { get; init; }
    public double Cortesia { get; init; }
    public double CtaCte { get; init; }
    public double Costo { get; init; }
    public string TTipoPedido { get; init; } = string.Empty;
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_RankingIntegrado</c>.
/// Legacy: <c>frmRepRankingIntegrado.frm</c>
/// Regla: BR-REP-008
/// </summary>
public sealed class RankingIntegradoRow
{
    public string TLocal { get; init; } = string.Empty;
    public string Local { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public double Cantidad { get; init; }
    public double Venta { get; init; }
    public double Neto { get; init; }
    public double Costo { get; init; }
    public double Produccion { get; init; }
    public double Cortesia { get; init; }
    public double CtaCte { get; init; }
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_VentaMensualIntegrado</c>.
/// Legacy: <c>frmRepVentaMensualIntegrado.frm</c>, <c>frmRepDiarioVentaIntegrado.frm</c>
/// Regla: BR-REP-009
/// </summary>
public sealed class VentaMensualIntegradoRow
{
    public int Dia { get; init; }
    public double Salon { get; init; }
    public double Delivery { get; init; }
    public double Llevar { get; init; }
    public double Canal4 { get; init; }
    public double Canal5 { get; init; }
    public double Venta { get; init; }
    public int Cantidad { get; init; }
    public int Pax { get; init; }
    public string TTipoPedido { get; init; } = string.Empty;
    public double Costo { get; init; }
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_AnaliticoMotorizadoIntegrado</c>.
/// Legacy: <c>frmRepAnaliticoMotorizadoIntegrado.frm</c>
/// Regla: BR-REP-010
/// </summary>
public sealed class AnaliticoMotorizadoIntegradoRow
{
    public string Motorizado { get; init; } = string.Empty;
    public string TipoProducto { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public double Cantidad { get; init; }
    public double Venta { get; init; }
    public double NPedidos { get; init; }
    public string TCodigoPedido { get; init; } = string.Empty;
    public double Comision { get; init; }
}

/// <summary>
/// Fila del dataset retornado por <c>USP_KDS_ResporteTiempoPedido</c>.
/// Regla: BR-REP-011
/// </summary>
public sealed class TiempoKdsPedidoRow
{
    public string Pedido { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public string TiempoCorto { get; init; } = string.Empty;
    public string TiempoLargo { get; init; } = string.Empty;
    public string TiempoPromedio { get; init; } = string.Empty;
}

/// <summary>
/// Fila del dataset retornado por <c>USP_KDS_ResporteTiempoProducto</c>.
/// Regla: BR-REP-012
/// </summary>
public sealed class TiempoKdsProductoRow
{
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public string TiempoCorto { get; init; } = string.Empty;
    public string TiempoLargo { get; init; } = string.Empty;
    public string TiempoPromedio { get; init; } = string.Empty;
}

/// <summary>
/// Envoltorio de resultado de reporte con metadatos para FastReport.
/// </summary>
public sealed class ReporteResultado<T>
{
    public IReadOnlyList<T> Filas { get; init; } = [];
    public string TituloReporte { get; init; } = string.Empty;
    public DateTime FechaEmision { get; init; } = DateTime.Now;
    public string NombrePlantilla { get; init; } = string.Empty;
}

/// <summary>
/// Expresión de precio configurable por el usuario.
/// Evita SQL injection al restringir a valores conocidos.
/// Regla: BR-REP-SQL-DYN-001 — nunca pasar string libre del usuario como @sPrecio.
/// </summary>
public enum ExpresionPrecio
{
    /// <summary>Precio de venta (<c>nVenta</c>)</summary>
    Venta,
    /// <summary>Precio neto (<c>nNeto</c>)</summary>
    Neto,
    /// <summary>Precio costo (<c>nInsumo+nGasto+nManoObra</c>)</summary>
    Costo
}

/// <summary>
/// Extensión para traducir <see cref="ExpresionPrecio"/> a la expresión SQL compatible con los SPs Legacy.
/// </summary>
public static class ExpresionPrecioExtensions
{
    /// <summary>
    /// Retorna la expresión SQL que el SP Legacy espera en el parámetro <c>@sPrecio</c>.
    /// </summary>
    public static string ToSqlExpresion(this ExpresionPrecio precio) => precio switch
    {
        ExpresionPrecio.Venta => "dbo.DPEDIDO.nVenta",
        ExpresionPrecio.Neto => "dbo.DPEDIDO.nNeto",
        ExpresionPrecio.Costo => "(dbo.DPEDIDO.nInsumo+dbo.DPEDIDO.nGasto+dbo.DPEDIDO.nManoObra)*dbo.DPEDIDO.nCantidad",
        _ => "dbo.DPEDIDO.nVenta"
    };
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_Anulacion</c>.
/// Legacy: <c>frmRepAnulado.frm</c> — "Control de Transacciones"
/// Regla: BR-REP-014
/// </summary>
public sealed class AnulacionRow
{
    public string TCodigoPedido { get; init; } = string.Empty;
    public string TItem { get; init; } = string.Empty;
    public string TCodigoProducto { get; init; } = string.Empty;
    public double NCantidad { get; init; }
    public double NVenta { get; init; }
    public string TEstadoItem { get; init; } = string.Empty;
    public string TDocumento { get; init; } = string.Empty;
    public DateTime FRegistro { get; init; }
    public bool LImprime { get; init; }
    public string TMotivoAnulacion { get; init; } = string.Empty;
    public string TObservacionAnulado { get; init; } = string.Empty;
    public string TUsuarioAnulado { get; init; } = string.Empty;
    public DateTime? FRegAnulado { get; init; }
    public string TTurno { get; init; } = string.Empty;
    public DateTime? FFechaItem { get; init; }
}

/// <summary>
/// Fila agregada retornada por <c>spRep_LiquidacionSuma</c>.
/// Legacy: <c>frmRepLiquidacionTicket.frm</c> — "Liquidación de Cajero por Ticketera"
/// Regla: BR-REP-015
/// </summary>
public sealed class LiquidacionTicketRow
{
    public double NNeto { get; init; }
    public double NImpuesto1 { get; init; }
    public double NImpuesto2 { get; init; }
    public double NImpuesto3 { get; init; }
    public double NVenta { get; init; }
    public string TTipoPedido { get; init; } = string.Empty;
    public int NTotalPromedio { get; init; }
    public int Total00 { get; init; }
}

/// <summary>
/// Fila del reporte "Cierre de Cajeros Delivery".
/// Legacy: <c>frmRepDeliveryTicket.frm</c> (query dinámica sobre MDOCUMENTO/DPREPAGO/MPEDIDO).
/// Regla: BR-REP-017
/// </summary>
public sealed class DeliveryTicketRow
{
    public string TCaja { get; init; } = string.Empty;
    public string TTipoPago { get; init; } = string.Empty;
    public string TipoPago { get; init; } = string.Empty;
    public string TMotorizado { get; init; } = string.Empty;
    public string Motorizado { get; init; } = string.Empty;
    public string TDocumento { get; init; } = string.Empty;
    public DateTime FRegistro { get; init; }
    public double NVenta { get; init; }
    public string TTurno { get; init; } = string.Empty;
    public string TUsuario { get; init; } = string.Empty;
    public string TMoneda { get; init; } = string.Empty;
    public string Mon { get; init; } = string.Empty;
    public double NTipoCambio { get; init; }
    public double NMonto { get; init; }
    public double NVuelto { get; init; }
    public string Tarjeta { get; init; } = string.Empty;
    public string TNumero { get; init; } = string.Empty;
    public string OtroTipo { get; init; } = string.Empty;
}

/// <summary>
/// Parámetros de filtro para el reporte "Cierre de Cajeros Delivery".
/// Legacy: <c>frmRepDeliveryTicket.frm</c> — filtros: turno o rango fechas, caja, motorizado.
/// Regla: BR-REP-017
/// </summary>
public sealed class DeliveryTicketParametros
{
    /// <summary>true = filtra por rango de fechas; false = turno específico.</summary>
    public bool TodosTurnos { get; init; } = true;
    public string Turno { get; init; } = string.Empty;
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }

    /// <summary>true = todas las cajas; false = sólo <see cref="Caja"/>.</summary>
    public bool TodasLasCajas { get; init; } = true;
    public string Caja { get; init; } = string.Empty;

    /// <summary>true = todos los motorizados; false = sólo <see cref="Motorizado"/>.</summary>
    public bool TodosLosMotorizados { get; init; } = true;
    public string Motorizado { get; init; } = string.Empty;
}

/// <summary>
/// Fila agregada del paloteo por ticketera.
/// Legacy: <c>frmRepPaloteoTicket.frm</c> (query dinámica sobre MPEDIDO/DPEDIDO/CPEDIDO).
/// Regla: BR-REP-016
/// </summary>
public sealed class PaloteoTicketRow
{
    public string TCodProducto { get; init; } = string.Empty;
    public string TLocal { get; init; } = string.Empty;
    public string Local { get; init; } = string.Empty;
    public string Salon { get; init; } = string.Empty;
    public string TipoProducto { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public double Cantidad { get; init; }
    public double Venta { get; init; }
}

// ── BR-REP-018 — Reporte de Reservas ─────────────────────────────────────────

/// <summary>
/// Criterio de ordenamiento del reporte de reservas.
/// Legacy: <c>frmRepReservas.frm</c> — <c>cboCriterio</c>
/// Regla: BR-REP-018
/// </summary>
public enum OrdenReserva
{
    Reserva = 0,
    Nombre = 1,
    Telefono = 2,
    Fecha = 3,
    Pax = 4,
    Estado = 5
}

/// <summary>
/// Fila del dataset retornado por la consulta dinámica de reservas.
/// Legacy: <c>frmRepReservas.frm</c> — Sub ObtenerReservas (TRESERVA + vEstadoReserva)
/// Regla: BR-REP-018
/// </summary>
public sealed class ReservaReporteRow
{
    public string TReserva { get; init; } = string.Empty;
    public DateTime FFecha { get; init; }
    public string TApellido { get; init; } = string.Empty;
    public string Cliente { get; init; } = string.Empty;
    public string TNombre { get; init; } = string.Empty;
    public string TTelefono { get; init; } = string.Empty;
    public int NPax { get; init; }
    public string TEstadoReserva { get; init; } = string.Empty;
    public string TObservacion { get; init; } = string.Empty;
    public DateTime FRegistro { get; init; }
    public string EstadoReserva { get; init; } = string.Empty;
}

/// <summary>
/// Parámetros de filtro para el reporte de reservas.
/// Legacy: <c>frmRepReservas.frm</c> — dtpFecIni/Fin + dtpHorIni/Fin + chkEstado + cboCriterio
/// Regla: BR-REP-018
/// </summary>
public sealed class ReservaReporteParametros
{
    public DateTime FechaHoraInicio { get; init; }
    public DateTime FechaHoraFin { get; init; }
    /// <summary>true = incluir estado '01' (Generado)</summary>
    public bool EstadoGenerado { get; init; } = true;
    /// <summary>true = incluir estado '02' (Atendido)</summary>
    public bool EstadoAtendido { get; init; }
    /// <summary>true = incluir estado '03' (Anulado)</summary>
    public bool EstadoAnulado { get; init; }
    public OrdenReserva Orden { get; init; } = OrdenReserva.Reserva;
}

// ── BR-REP-019 — Reporte de Entregas ─────────────────────────────────────────

/// <summary>
/// Formatos del reporte de entregas.
/// Legacy: <c>frmRepEntrega.frm</c> (Option1/Option2/Option3)
/// </summary>
public enum FormatoReporteEntrega
{
    DetalladoFormato1 = 0,
    ResumidoPorProducto = 1,
    DetalladoFormato2 = 2
}

/// <summary>
/// Fila del dataset retornado por <c>spRep_Entregas</c>.
/// Legacy: <c>frmRepEntrega.frm</c>
/// Regla: BR-REP-019
/// </summary>
public sealed class EntregaRow
{
    public string Pedido { get; init; } = string.Empty;
    public string FechaPedido { get; init; } = string.Empty;
    public string FechaEntrega { get; init; } = string.Empty;
    public string HoraEntrega { get; init; } = string.Empty;
    public string Identificacion { get; init; } = string.Empty;
    public string Cliente { get; init; } = string.Empty;
    public string Direccion { get; init; } = string.Empty;
    public string Telefono { get; init; } = string.Empty;
    public string Observacion { get; init; } = string.Empty;
    public string CanalVenta { get; init; } = string.Empty;
    public string EstadoPedido { get; init; } = string.Empty;
    public string Cancelacion { get; init; } = string.Empty;
    public string Producto { get; init; } = string.Empty;
    public double Cantidad { get; init; }
    public double Monto { get; init; }
    public double SaldoPendiente { get; init; }
    public string DObservacion { get; init; } = string.Empty;
    public string DPropiedad { get; init; } = string.Empty;
}

/// <summary>
/// Parámetros de filtro para <c>spRep_Entregas</c>.
/// Legacy: <c>frmRepEntrega.frm</c>
/// Regla: BR-REP-019
/// </summary>
public sealed class EntregaParametros
{
    public DateTime FechaHoraInicio { get; init; }
    public DateTime FechaHoraFin { get; init; }
    public string CodigoCliente { get; init; } = string.Empty;
    public string EstadoEntrega { get; init; } = string.Empty;
    public string EstadoCancelacion { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string CodigoProducto { get; init; } = string.Empty;
    public string CanalVenta { get; init; } = string.Empty;
    public FormatoReporteEntrega Formato { get; init; } = FormatoReporteEntrega.ResumidoPorProducto;
}

// ── BR-REP-020 — Venta Mensual por Fechas ─────────────────────────────────────

/// <summary>
/// Tipo de precio utilizado para el cálculo de venta.
/// Legacy: <c>frmRepVentaFecha.frm</c> — optValor (0=Precio venta / 1=Precio neto)
/// Regla: BR-REP-020
/// </summary>
public enum TipoPrecioVentaFecha
{
    /// <summary>nVenta del detalle de pedido (precio de venta).</summary>
    Venta = 0,
    /// <summary>nPrecioNeto * nCantidad del detalle de pedido (precio neto).</summary>
    Neto = 1
}

/// <summary>
/// Fila agregada por día retornada por <c>spRep_VentaFecha</c>.
/// Legacy: <c>frmRepVentaFecha.frm</c> — columnas Dia/Salon/Delivery/Llevar/Canal4/Canal5/Venta/Cantidad/Pax/Fecha
/// Regla: BR-REP-020
/// </summary>
public sealed class VentaFechaRow
{
    public int Dia { get; init; }
    public double Salon { get; init; }
    public double Delivery { get; init; }
    public double Llevar { get; init; }
    public double Canal4 { get; init; }
    public double Canal5 { get; init; }
    public double Venta { get; init; }
    public int Cantidad { get; init; }
    public int Pax { get; init; }
    public DateTime Fecha { get; init; }
}

/// <summary>
/// Parámetros de filtro para <c>spRep_VentaFecha</c>.
/// Legacy: <c>frmRepVentaFecha.frm</c> — dtpAnual/CmbMes/dtpHora/optValor/chkCFacturados/ChkDocumento/OptSel/grdSubGrupos
/// Regla: BR-REP-020
/// </summary>
public sealed class VentaFechaParametros
{
    /// <summary>Año del mes a reportar.</summary>
    public int Ano { get; init; }
    /// <summary>Mes del año (1–12).</summary>
    public int Mes { get; init; }
    /// <summary>Hora de corte para asignación de día contable (0–23). Legacy: dtpHora.</summary>
    public int HoraCierre { get; init; }
    /// <summary>Tipo de precio: Venta o Neto.</summary>
    public TipoPrecioVentaFecha TipoPrecio { get; init; } = TipoPrecioVentaFecha.Venta;
    /// <summary>Si true, valoriza con cero la pre-venta facturada (chkCFacturados).</summary>
    public bool ValorarPreventaEnCero { get; init; }
    /// <summary>Si true, evalúa la venta por documentos emitidos (tipooper=2). Legacy: ChkDocumento.</summary>
    public bool EvaluarPorDocumentos { get; init; }
    /// <summary>
    /// Códigos de sub-grupo para filtrar. Vacío = todos (OptSel(0)="Todos").
    /// Cuando contiene elementos, aplica filtro tCodigoSubGrupo IN (...).
    /// </summary>
    public IReadOnlyList<string> SubGruposFiltro { get; init; } = Array.Empty<string>();
}

/// <summary>
/// Ítem de sub-grupo para la UI de filtros.
/// Legacy: <c>vSubGrupo</c> — columns Codigo/Descripcion/tGrupo
/// </summary>
public sealed class SubGrupoItem
{
    public string Codigo { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
}

// ============================================================
// Liquidación de Cajero — frmRepLiquidacion.frm
// Legacy: spRep_LiquidacionOutPut, spRep_Liquidacion, spRep_LiquidacionSuma
// Regla: BR-REP-021
// ============================================================

/// <summary>
/// Modo de filtro temporal: por turno específico o por rango de fechas.
/// Legacy: ChkTurno.Value — 0=Turno, 1=FechaRango
/// </summary>
public enum LiquidacionModoFiltro { PorTurno = 0, PorFecha = 1 }

/// <summary>
/// Parámetros de filtro para el reporte Liquidación de Cajero.
/// Legacy: <c>frmRepLiquidacion.frm</c>
/// Regla: BR-REP-021
/// </summary>
public sealed class LiquidacionParametros
{
    /// <summary>Modo de filtro temporal (Turno o Rango de Fechas). Legacy: ChkTurno.</summary>
    public LiquidacionModoFiltro ModoFiltro { get; init; } = LiquidacionModoFiltro.PorFecha;
    /// <summary>Código de turno. Requerido cuando ModoFiltro = PorTurno. Legacy: sTurno.</summary>
    public string Turno { get; init; } = string.Empty;
    /// <summary>Fecha/hora inicio. Legacy: dtpFecIni + dtpHorIni.</summary>
    public DateTime FechaInicio { get; init; }
    /// <summary>Fecha/hora fin. Legacy: dtpFecFin + dtpHorFin.</summary>
    public DateTime FechaFin { get; init; }
    /// <summary>Si true, filtra por día contable (solo fecha, sin hora). Legacy: chkDiaContable.</summary>
    public bool PorDiaContable { get; init; }
    /// <summary>Código de usuario. Vacío = todos. Legacy: cboUsuario (chkUsuario).</summary>
    public string Usuario { get; init; } = string.Empty;
    /// <summary>Código de sector de venta. Vacío = todos. Legacy: cboSectorVenta (chkSectorVenta).</summary>
    public string SectorVenta { get; init; } = string.Empty;
    /// <summary>Si true, incluye cortesías en el cálculo. Legacy: chkCortesia.</summary>
    public bool IncluirCortesia { get; init; }
    /// <summary>Si true, muestra todos los tipos de documentos/pagos (chkGenerado = 1). Legacy: chkGenerado.</summary>
    public bool MostrarTodos { get; init; } = true;
    /// <summary>Si true, usa SPs variante _NC (nota crédito Ofisis). Legacy: lNcOfisis.</summary>
    public bool UsarVarianteNc { get; init; }
}

/// <summary>
/// Totales escalares devueltos por <c>spRep_LiquidacionOutPut</c> (parámetros OUTPUT).
/// Legacy: nDolar, nNeto, nImpuesto1..3, nVenta, nDescuento, nRecargo, nCambio, nAdulto..5, nNino..5
/// </summary>
public sealed class LiquidacionOutput
{
    public double Dolar { get; init; }
    public double DolarCambio { get; init; }
    public double OtroDoc { get; init; }
    public double NoCobrado { get; init; }
    public double Neto { get; init; }
    public double Impuesto1 { get; init; }
    public double Impuesto2 { get; init; }
    public double Impuesto3 { get; init; }
    public double VentaTotal { get; init; }
    public double Descuento { get; init; }
    public double Recargo { get; init; }
    public double TipoCambio { get; init; }
    public double Adulto { get; init; }
    public double Nino { get; init; }
    public double Adulto2 { get; init; }
    public double Nino2 { get; init; }
    public double Adulto3 { get; init; }
    public double Nino3 { get; init; }
    public double Adulto4 { get; init; }
    public double Nino4 { get; init; }
    public double Adulto5 { get; init; }
    public double Nino5 { get; init; }
    public double PagadosEnOtroTurno { get; init; }
}

/// <summary>
/// Fila del dataset principal retornado por <c>spRep_Liquidacion</c> @flagTipo='1' (documentos).
/// Campos: tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago,
///         nTC, nVenta, nVenta1, nVenta2, tObservacion, tTipoDocumento, tInicio, tFinal,
///         tEmitido, tAnulado, total00, total14, totalNF, nVentaME2, nVentaME3, nVuelto
/// Legacy: dsrLiquidacion dataset
/// </summary>
public sealed class LiquidacionRow
{
    public string TGrupo { get; init; } = string.Empty;
    public string Grupo { get; init; } = string.Empty;
    public string TSubGrupo { get; init; } = string.Empty;
    public string SubGrupo { get; init; } = string.Empty;
    public string TDocumento { get; init; } = string.Empty;
    public string TUsuario { get; init; } = string.Empty;
    public string FFecha { get; init; } = string.Empty;
    public string FPago { get; init; } = string.Empty;
    public double NTC { get; init; }
    public double NVenta { get; init; }
    public double NVenta1 { get; init; }
    public double NVenta2 { get; init; }
    public string TObservacion { get; init; } = string.Empty;
    public string TTipoDocumento { get; init; } = string.Empty;
    public string TInicio { get; init; } = string.Empty;
    public string TFinal { get; init; } = string.Empty;
    public string TEmitido { get; init; } = string.Empty;
    public string TAnulado { get; init; } = string.Empty;
    public double Total00 { get; init; }
    public double Total14 { get; init; }
    public double TotalNF { get; init; }
    public double NVentaME2 { get; init; }
    public double NVentaME3 { get; init; }
    public double NVuelto { get; init; }
}

/// <summary>
/// Fila de sumas por grupo de pago retornada por <c>spRep_Liquidacion</c> @flagTipo='2'.
/// Campos: tGrupo, nVenta1 (MN), nVenta2 (ME), nVenta3 (ME equivalente MN)
/// Legacy: RsSumas filtros tGrupo 01=Efectivo, 03=Cheque, 04=Otro, 05=Puntos, 06=Cortesía, 07=CtaCobrar, 08=Ingreso, 09=IngresoAnticipo, 10=Egreso, 20=NotaCredito
/// </summary>
public sealed class LiquidacionSumaGrupoRow
{
    public string TGrupo { get; init; } = string.Empty;
    public double NVenta1 { get; init; }
    public double NVenta2 { get; init; }
    public double NVenta3 { get; init; }
}

/// <summary>
/// Fila de tarjeta de crédito retornada por <c>spRep_Liquidacion</c> @flagTipo='3'.
/// Campos: tSubGrupo (01-24), nVenta1 (importe tarjeta), nVenta2 (propina)
/// Legacy: RsSumas iteración i=1..24 tSubGrupo
/// </summary>
public sealed class LiquidacionTarjetaRow
{
    public string TSubGrupo { get; init; } = string.Empty;
    public double NVenta1 { get; init; }
    public double NVenta2 { get; init; }
}

/// <summary>
/// Fila por tipo de pedido retornada por <c>spRep_LiquidacionSuma</c>.
/// Campos: ttipopedido, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nTotalPromedio, total00
/// Legacy: RsSumas filtros ttipopedido 01=Salón, 02=Delivery, 03=Llevar, 04=Canal4, 05=Canal5, XX=Fiscal
/// </summary>
public sealed class LiquidacionTipoPedidoRow
{
    public string TTipoPedido { get; init; } = string.Empty;
    public double NNeto { get; init; }
    public double NImpuesto1 { get; init; }
    public double NImpuesto2 { get; init; }
    public double NImpuesto3 { get; init; }
    public double NVenta { get; init; }
    public double NTotalPromedio { get; init; }
    public double Total00 { get; init; }
}

/// <summary>
/// Fila de otros tipos de cancelación retornada por <c>spRep_Liquidacion</c> @flagTipo='5'.
/// Campos: CODIGO, Descripcion, nVenta1, nVenta2, nVenta3 (MN/ME/MEenMN)
/// Legacy: RsSumas + vTipoCancelacion iteración i=1..20
/// </summary>
public sealed class LiquidacionOtroTipoRow
{
    public string Codigo { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public double NVenta1 { get; init; }
    public double NVenta2 { get; init; }
    public double NVenta3 { get; init; }
}

/// <summary>
/// Filas de la vista <c>vIngreso</c> por tarjeta para recibos de ingreso/tarjeta.
/// Legacy: RsTarjetas — select tTarjeta, sum(nMonto) nVenta1 From vIngreso where ... group by tTarjeta
/// </summary>
public sealed class LiquidacionIngresoTarjetaRow
{
    public string TarjetaCodigo { get; init; } = string.Empty;
    public double Monto { get; init; }
}

/// <summary>
/// Resultado completo del reporte Liquidación de Cajero.
/// Agrega todos los conjuntos de datos necesarios para renderizar el reporte.
/// Legacy: dsrLiquidacion / dsrLiquidacionVenta
/// Regla: BR-REP-021
/// </summary>
public sealed class LiquidacionResultado
{
    /// <summary>Filas de documentos (spRep_Liquidacion tipo 1).</summary>
    public IReadOnlyList<LiquidacionRow> Documentos { get; init; } = Array.Empty<LiquidacionRow>();
    /// <summary>Sumas por grupo de pago (spRep_Liquidacion tipo 2).</summary>
    public IReadOnlyList<LiquidacionSumaGrupoRow> SumasGrupo { get; init; } = Array.Empty<LiquidacionSumaGrupoRow>();
    /// <summary>Sumas por tarjeta de crédito (spRep_Liquidacion tipo 3).</summary>
    public IReadOnlyList<LiquidacionTarjetaRow> Tarjetas { get; init; } = Array.Empty<LiquidacionTarjetaRow>();
    /// <summary>Sumas por tipo de pedido (spRep_LiquidacionSuma).</summary>
    public IReadOnlyList<LiquidacionTipoPedidoRow> TiposPedido { get; init; } = Array.Empty<LiquidacionTipoPedidoRow>();
    /// <summary>Otros tipos de cancelación (spRep_Liquidacion tipo 5).</summary>
    public IReadOnlyList<LiquidacionOtroTipoRow> OtrosTipos { get; init; } = Array.Empty<LiquidacionOtroTipoRow>();
    /// <summary>Totales escalares del SP OUTPUT.</summary>
    public LiquidacionOutput Output { get; init; } = new();
    /// <summary>Título descriptivo del reporte (turno/fechas/usuario).</summary>
    public string Titulo { get; init; } = string.Empty;
}
