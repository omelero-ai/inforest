namespace Inforest.Application.Delivery;

// ──────────────────────────────────────────────────────────────────────────────
// DTOs — Búsqueda de Clientes Delivery (frmBusquedaDelivery.frm — POS-FUNC-036)
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Ítem de la grilla principal de búsqueda de clientes delivery.
/// <para>
/// Legacy: <c>frmBusquedaDelivery.frm</c> Form_Load — query principal sobre
/// <c>TDELIVERY LEFT JOIN vZona</c>.
/// Reglas: BR-DEL-036.
/// </para>
/// </summary>
public sealed record ClienteDeliveryBusquedaItem(
    string CodigoDelivery,
    string? Telefono,
    string Cliente,
    decimal Descuento,
    string? Referencia,
    string? Zona,
    decimal Acumulado,
    decimal Utilizado,
    decimal Disponible);

/// <summary>
/// Ítem de la grilla principal de mantenimiento de clientes delivery.
/// <para>
/// Legacy: <c>frmClienteDelivery.frm</c> Form_Load —
/// <c>SELECT *, str(nDescuento,10,2) as xDescuento, str(nLinea,10,2) as xLinea FROM vDelivery</c>.
/// Reglas: BR-DEL-037.
/// </para>
/// </summary>
public sealed record ClienteDeliveryListadoItem(
    string Codigo,
    string? TipoCliente,
    string? Cliente,
    string? Telefono,
    string? Direccion,
    decimal Linea,
    decimal Descuento,
    string? EstadoFrecuente,
    bool Activo);

/// <summary>
/// Detalle completo de un cliente delivery para el panel lateral del formulario.
/// <para>
/// Legacy: <c>frmBusquedaDelivery.frm</c> Sub Asigna() — <c>SELECT * FROM vDelivery WHERE Codigo=…</c>.
/// Reglas: BR-DEL-036.
/// </para>
/// </summary>
public sealed record ClienteDeliveryDetalleBusqueda(
    string CodigoDelivery,
    string? Apellido,
    string? Nombre,
    string? Telefono,
    string? Direccion,
    string? Zona,
    string? Referencia,
    string? Observacion,
    decimal Descuento,
    decimal Acumulado,
    decimal Utilizado,
    decimal Disponible,
    string? TipoCliente);

/// <summary>
/// Estadísticas históricas de un cliente delivery ("Otros Datos").
/// <para>
/// Legacy: <c>frmBusquedaDelivery.frm</c> cmdOpcion(3) "Otros Datos" —
/// COUNT/SUM sobre <c>MPEDIDO/DPEDIDO</c> + último pedido sobre <c>MDOCUMENTO</c>.
/// Reglas: BR-DEL-036.
/// </para>
/// </summary>
public sealed record EstadisticasClienteDelivery(
    /// <summary>Fecha desde la que se calculan las estadísticas (servidor - nDiasDelivery).</summary>
    DateOnly? Desde,
    /// <summary>Número de pedidos activos del cliente en el período.</summary>
    int NumeroPedidos,
    /// <summary>Suma de ventas activas del cliente en el período.</summary>
    decimal TotalVenta,
    /// <summary>Fecha del último documento emitido para el cliente.</summary>
    DateTime? FechaUltimaCompra,
    /// <summary>Código del último documento emitido.</summary>
    string? UltimoDocumento,
    /// <summary>Monto del último documento emitido.</summary>
    decimal MontoUltimoDoc,
    /// <summary>Razón social del cliente fiscal del último documento.</summary>
    string? RazonSocial,
    /// <summary>RUC/identidad del cliente fiscal del último documento.</summary>
    string? Ruc);

/// <summary>
/// Vista de una tienda (sucursal) de un cliente delivery.
/// <para>
/// Legacy: <c>frmBusquedaDelivery.frm</c> Tienda_Click — <c>SELECT * FROM vTienda WHERE lActivo=1 AND tCodigoDelivery=…</c>.
/// </para>
/// </summary>
public sealed record TiendaDeliveryItem(
    string CodigoTienda,
    string Descripcion,
    string? Direccion);
