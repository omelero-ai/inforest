namespace Inforest.Domain.Entities.Hardware;

/// <summary>
/// Modelo de ticket para BlueVision TVS.
/// Legacy: modBlueVision.bas — TvsTicket (id, mode, room, table, total, TPV, evento, mozo).
/// Regla BR-PERIPH-003.
/// </summary>
public sealed class BlueVisionTicket
{
    public string Id { get; }
    public string Area { get; }
    public string Mesa { get; }
    public string Mozo { get; }
    public string Evento { get; }
    public decimal Total { get; }
    public DateTime Timestamp { get; }
    public IReadOnlyList<BlueVisionLinea> Lineas { get; }

    private BlueVisionTicket(
        string id, string area, string mesa, string mozo,
        string evento, decimal total, DateTime timestamp,
        IReadOnlyList<BlueVisionLinea> lineas)
    {
        Id = id;
        Area = area;
        Mesa = mesa;
        Mozo = mozo;
        Evento = evento;
        Total = total;
        Timestamp = timestamp;
        Lineas = lineas;
    }

    public static BlueVisionTicket Crear(
        string id, string area, string mesa, string mozo,
        string evento, decimal total, IReadOnlyList<BlueVisionLinea> lineas)
    {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Id requerido.", nameof(id));
        if (string.IsNullOrWhiteSpace(area)) throw new ArgumentException("Area requerida.", nameof(area));

        return new BlueVisionTicket(id, area, mesa, mozo ?? "SIN MOZO",
            evento, total, DateTime.Now, lineas);
    }
}

/// <summary>
/// Línea de ticket BlueVision.
/// Legacy: TvsLine — concept, id, mode, price, room, table, TPV, waiter, ticketId, total, units, timestamp.
/// </summary>
public sealed class BlueVisionLinea
{
    public string Id { get; }
    public string TicketId { get; }
    public string Concepto { get; }
    public decimal Precio { get; }
    public decimal Cantidad { get; }
    public decimal Total { get; }

    private BlueVisionLinea(string id, string ticketId, string concepto, decimal precio, decimal cantidad, decimal total)
    {
        Id = id;
        TicketId = ticketId;
        Concepto = concepto;
        Precio = precio;
        Cantidad = cantidad;
        Total = total;
    }

    public static BlueVisionLinea Crear(string id, string ticketId, string concepto, decimal precio, decimal cantidad)
    {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Id requerido.", nameof(id));
        return new BlueVisionLinea(id, ticketId, concepto, precio, cantidad, precio * cantidad);
    }
}
