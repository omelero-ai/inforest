namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Opciones de configuración de BlueVision TVS.
/// Legacy: modBlueVision.bas — BLUEVISION.INI (login, ClearPassword, Url).
/// ADR-012: se migra de INI a appsettings.json.
/// </summary>
public sealed class BlueVisionOptions
{
    public const string SectionName = "Hardware:BlueVision";

    /// <summary>Habilita el servicio BlueVision. Si false, se usa NullBlueVisionService.</summary>
    public bool Habilitado { get; set; }

    /// <summary>URL del servidor BlueVision TVS. Legacy: BLUEVISION.INI [BlueVision] Url.</summary>
    public string Url { get; set; } = "http://192.168.3.86:81/";

    /// <summary>Usuario del servidor BlueVision. Legacy: BLUEVISION.INI [BlueVision] login.</summary>
    public string Login { get; set; } = string.Empty;

    /// <summary>Contraseña en texto plano del servidor BlueVision. Legacy: BLUEVISION.INI [BlueVision] ClearPassword.</summary>
    public string ClearPassword { get; set; } = string.Empty;
}
