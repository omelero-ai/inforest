namespace Inforest.Application.Interfaces;

/// <summary>
/// Abstracción para el servicio de impresión.
/// Legacy: modProcedimiento.bas — ImprimirTicket(), impresoras ESC/POS.
/// Hardware: Impresoras térmicas (ESC/POS), impresoras por área (cocina, bar, etc.)
/// </summary>
public interface IImpresoraService
{
    /// <summary>
    /// Imprime un ticket/recibo de venta.
    /// Legacy: Múltiples rutinas de impresión en modProcedimiento.bas
    /// </summary>
    Task ImprimirTicketAsync(string contenido, string? nombreImpresora = null, CancellationToken cancellationToken = default);

    /// <summary>
    /// Imprime una comanda a un área específica (cocina, bar, etc.)
    /// Legacy: ImprimirComanda(), TAREA — impresoras por área.
    /// </summary>
    Task ImprimirComandaAsync(string contenido, string codigoArea, CancellationToken cancellationToken = default);

    /// <summary>
    /// Abre el cajón de dinero.
    /// Legacy: AbrirCajon() en modProcedimiento.bas
    /// </summary>
    Task AbrirCajonAsync(string? nombreImpresora = null, CancellationToken cancellationToken = default);
}
