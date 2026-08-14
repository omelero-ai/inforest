using System.Net.Http.Json;
using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Adaptador BlueVision TVS via HttpClient nativo .NET 8.
/// Legacy: modBlueVision.bas — TvsClientApi (COM) con CreateSession, SaveTicket, SaveTicketLine.
/// Regla BR-PERIPH-003. ADR-012.
/// </summary>
internal sealed class BlueVisionHttpClient : IBlueVisionService
{
    private readonly HttpClient _httpClient;
    private readonly BlueVisionOptions _options;
    private readonly ILogger<BlueVisionHttpClient> _logger;

    public BlueVisionHttpClient(
        HttpClient httpClient,
        IOptions<BlueVisionOptions> options,
        ILogger<BlueVisionHttpClient> logger)
    {
        _httpClient = httpClient;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<bool> EnviarTicketAsync(BlueVisionTicketDto ticket, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("BlueVision: enviando ticket {TicketId} mesa={Mesa} area={Area}", ticket.Id, ticket.Mesa, ticket.Area);
        try
        {
            // 1 — Crear sesión (equivale a TvsClientApi.CreateSession)
            var sessionPayload = new { login = _options.Login, clearPassword = _options.ClearPassword };
            var sessionResp = await _httpClient.PostAsJsonAsync("api/session", sessionPayload, cancellationToken);
            if (!sessionResp.IsSuccessStatusCode)
            {
                _logger.LogWarning("BlueVision: fallo al crear sesión, status={Status}", sessionResp.StatusCode);
                return false;
            }

            // 2 — Enviar cabecera del ticket (SaveTicket)
            var ticketPayload = new
            {
                id = ticket.Id,
                mode = "--",
                room = ticket.Area,
                table = ticket.Mesa,
                ticketNumber = 1,
                total = ticket.Total,
                tpv = ticket.Area,
                tvsEventName = ticket.Evento,
                waiter = ticket.Mozo,
                timeStamp = DateTime.Now
            };
            var ticketResp = await _httpClient.PostAsJsonAsync("api/ticket", ticketPayload, cancellationToken);
            if (!ticketResp.IsSuccessStatusCode)
            {
                _logger.LogWarning("BlueVision: fallo al guardar ticket {TicketId}, status={Status}", ticket.Id, ticketResp.StatusCode);
                return false;
            }

            // 3 — Enviar líneas del ticket (SaveTicketLine)
            foreach (var linea in ticket.Lineas)
            {
                var lineaPayload = new
                {
                    id = linea.Id,
                    ticketId = ticket.Id,
                    concept = linea.Concepto,
                    mode = "--",
                    price = linea.Precio,
                    room = ticket.Area,
                    table = ticket.Mesa,
                    tpv = ticket.Area,
                    waiter = ticket.Mozo,
                    total = linea.Total,
                    units = linea.Cantidad,
                    timeStamp = DateTime.Now
                };
                var lineaResp = await _httpClient.PostAsJsonAsync("api/ticket/line", lineaPayload, cancellationToken);
                if (!lineaResp.IsSuccessStatusCode)
                    _logger.LogWarning("BlueVision: fallo al guardar línea {LineaId}, status={Status}", linea.Id, lineaResp.StatusCode);
            }

            _logger.LogInformation("BlueVision: ticket {TicketId} enviado correctamente", ticket.Id);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "BlueVision: error al enviar ticket {TicketId}", ticket.Id);
            return false;
        }
    }
}
