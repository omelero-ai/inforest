namespace Inforest.Application.Interfaces;

/// <summary>
/// Servicio de integración con Rappi (órdenes externas).
/// <para>
/// Legacy: <c>frmOrdenesConsola.frm</c>, <c>frmEnvioDatosRappi.frm</c>.
/// Habilitado por flag <c>lOrdenesRappi</c> en <c>TCAJA</c>.
/// </para>
/// Regla BR-DEL-004.
/// </summary>
public interface IRappiOrderService
{
    /// <summary>
    /// Obtiene órdenes externas pendientes de procesar desde Rappi.
    /// Legacy: polling en <c>frmOrdenesConsola</c>.
    /// </summary>
    Task<IEnumerable<OrdenExterna>> ObtenerOrdenesPendientesAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Envía datos de una orden al sistema Rappi.
    /// Legacy: <c>frmEnvioDatosRappi</c> — serialización y envío.
    /// </summary>
    Task<bool> EnviarOrdenAsync(string codigoPedido, CancellationToken cancellationToken = default);

    /// <summary>
    /// Indica si la integración Rappi está habilitada para la caja actual.
    /// Legacy: flag <c>lOrdenesRappi</c> en <c>TCAJA</c>.
    /// Regla BR-DEL-004.
    /// </summary>
    bool EstaHabilitado { get; }
}

/// <summary>
/// Representación de una orden externa proveniente de Rappi u otro canal externo.
/// </summary>
public sealed record OrdenExterna(
    string IdOrdenExterna,
    string Origen,
    string? NombreCliente,
    string? Telefono,
    string? Direccion,
    decimal MontoTotal,
    DateTime FechaRecepcion,
    string EstadoExterno,
    string? RawPayload);
