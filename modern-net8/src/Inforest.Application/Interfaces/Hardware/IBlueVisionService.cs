namespace Inforest.Application.Interfaces.Hardware;

/// <summary>
/// Abstracción para el servicio BlueVision TVS (visualización de tickets en pantallas de mesa).
/// Legacy: modBlueVision.bas — BlueVision_Core_TVS.dll (TvsClientApi, CreateSession, SaveTicket, SaveTicketLine).
/// Regla BR-PERIPH-003. ADR-012.
/// </summary>
public interface IBlueVisionService
{
    /// <summary>
    /// Envía un ticket completo (cabecera + líneas) al servidor BlueVision TVS.
    /// Legacy: TVS_EnviarTicket(RsImpresion, sArea, sEvento, sPedido)
    /// </summary>
    Task<bool> EnviarTicketAsync(BlueVisionTicketDto ticket, CancellationToken cancellationToken = default);
}

/// <summary>
/// Datos del ticket para BlueVision TVS.
/// Legacy: TvsTicket — id, mode, room, table, total, TPV, evento, mozo, timestamp.
/// </summary>
public sealed record BlueVisionTicketDto(
    string Id,
    string Area,
    string Mesa,
    string Mozo,
    string Evento,
    decimal Total,
    IReadOnlyList<BlueVisionLineaDto> Lineas);

/// <summary>
/// Línea de ticket BlueVision.
/// Legacy: TvsLine — concept, id, mode, price, room, table, TPV, waiter, ticketId, total, units, timestamp.
/// </summary>
public sealed record BlueVisionLineaDto(
    string Id,
    string Concepto,
    decimal Precio,
    decimal Cantidad,
    decimal Total);
