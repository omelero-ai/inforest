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
