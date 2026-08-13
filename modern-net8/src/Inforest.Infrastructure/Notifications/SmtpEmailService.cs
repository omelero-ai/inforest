using System.Net;
using System.Net.Mail;
using Inforest.Application.Interfaces;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Inforest.Infrastructure.Notifications;

/// <summary>
/// Opciones de configuración SMTP.
/// Legacy: Chilkat COM component con SmtpHost/Port configurados en INI + TPARAMETRO.
/// </summary>
public sealed class SmtpOptions
{
    public const string SectionName = "Smtp";

    public string Host { get; init; } = string.Empty;
    public int Port { get; init; } = 25;
    public string Usuario { get; init; } = string.Empty;
    public string Clave { get; init; } = string.Empty;
    public string CorreoDesde { get; init; } = string.Empty;
    public string NombreDesde { get; init; } = string.Empty;

    /// <summary>Asunto/cuerpo fijo del correo de prórroga de licencia.</summary>
    public string AsuntoProrrogaLicencia { get; init; } = "Prórroga de licencia INFOREST";
    public string CuerpoProrrogaLicencia { get; init; } = "Se ha generado una solicitud de prórroga para la licencia del sistema INFOREST.";
}

/// <summary>
/// Servicio SMTP de notificaciones por correo electrónico.
/// Legacy: claCorreoElectronico.cls — EnviarCorreo / EnviarCorreoProrroga.
/// </summary>
internal sealed class SmtpEmailService : INotificacionEmailService
{
    private readonly SmtpOptions _options;
    private readonly ILogger<SmtpEmailService> _logger;

    public SmtpEmailService(IOptions<SmtpOptions> options, ILogger<SmtpEmailService> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    /// <inheritdoc/>
    public async Task<bool> EnviarCorreoAsync(
        string destinatario,
        string asunto,
        string cuerpo,
        string? adjuntoRuta = null,
        CancellationToken ct = default)
    {
        try
        {
            using var message = BuildMessage(destinatario, asunto, cuerpo, adjuntoRuta);
            using var client = BuildClient();
            await client.SendMailAsync(message, ct);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error enviando correo a {Destinatario}", destinatario);
            return false;
        }
    }

    /// <inheritdoc/>
    public async Task EnviarCorreoProrrogaAsync(string destinatario, CancellationToken ct = default)
    {
        await EnviarCorreoAsync(
            destinatario,
            _options.AsuntoProrrogaLicencia,
            _options.CuerpoProrrogaLicencia,
            null,
            ct);
    }

    private MailMessage BuildMessage(string destinatario, string asunto, string cuerpo, string? adjunto)
    {
        // Legacy: admitía múltiples destinatarios separados por ";" — se normaliza a "," (igual que el Legacy)
        var normalizado = destinatario.Replace(';', ',');

        var message = new MailMessage
        {
            From    = new MailAddress(_options.CorreoDesde, _options.NombreDesde),
            Subject = asunto,
            Body    = cuerpo
        };

        foreach (var addr in normalizado.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            message.To.Add(addr);

        if (!string.IsNullOrWhiteSpace(adjunto) && File.Exists(adjunto))
            message.Attachments.Add(new Attachment(adjunto));

        return message;
    }

    private SmtpClient BuildClient()
        => new(_options.Host, _options.Port)
        {
            Credentials = new NetworkCredential(_options.Usuario, _options.Clave),
            EnableSsl   = _options.Port == 587 || _options.Port == 465
        };
}
