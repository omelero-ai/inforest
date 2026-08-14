namespace Inforest.Application.Interfaces;

/// <summary>
/// Contrato para envío de correos electrónicos.
/// Legacy: claCorreoElectronico.cls — Chilkat COM component.
/// </summary>
public interface INotificacionEmailService
{
    /// <summary>
    /// Envía un correo electrónico.
    /// Legacy: EnviarCorreo() — claCorreoElectronico.cls.
    /// </summary>
    Task<bool> EnviarCorreoAsync(
        string destinatario,
        string asunto,
        string cuerpo,
        string? adjuntoRuta = null,
        CancellationToken ct = default);

    /// <summary>
    /// Envía un correo de prórroga de licencia.
    /// Legacy: EnviarCorreoProrroga() — claCorreoElectronico.cls.
    /// </summary>
    Task EnviarCorreoProrrogaAsync(string destinatario, CancellationToken ct = default);
}
