namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Opciones de configuración de BlueVision TVS.
/// Legacy: modBlueVision.bas — BLUEVISION.INI (login, ClearPassword, Url).
/// ADR-012: se migra de INI a configuración externalizada (ADR-005).
///
/// IMPORTANTE — seguridad (ADR-005):
/// Login y ClearPassword NO deben almacenarse en appsettings.json (archivo comprometido en repo).
/// Proveerlos vía variables de entorno:
///   INFOREST_Hardware__BlueVision__Login=&lt;valor&gt;
///   INFOREST_Hardware__BlueVision__ClearPassword=&lt;valor&gt;
/// o vía dotnet user-secrets en desarrollo:
///   dotnet user-secrets set "Hardware:BlueVision:Login" "&lt;valor&gt;"
///   dotnet user-secrets set "Hardware:BlueVision:ClearPassword" "&lt;valor&gt;"
/// </summary>
public sealed class BlueVisionOptions
{
    public const string SectionName = "Hardware:BlueVision";

    /// <summary>Habilita el servicio BlueVision. Si false, se usa NullBlueVisionService.</summary>
    public bool Habilitado { get; set; }

    /// <summary>URL del servidor BlueVision TVS. Legacy: BLUEVISION.INI [BlueVision] Url.</summary>
    public string Url { get; set; } = "http://192.168.3.86:81/";

    /// <summary>
    /// Usuario del servidor BlueVision. Legacy: BLUEVISION.INI [BlueVision] login.
    /// Debe configurarse vía variable de entorno INFOREST_Hardware__BlueVision__Login. Ver ADR-005.
    /// </summary>
    public string Login { get; set; } = string.Empty;

    /// <summary>
    /// Contraseña del servidor BlueVision. Legacy: BLUEVISION.INI [BlueVision] ClearPassword.
    /// Debe configurarse vía variable de entorno INFOREST_Hardware__BlueVision__ClearPassword. Ver ADR-005.
    /// NUNCA almacenar en appsettings.json.
    /// </summary>
    public string ClearPassword { get; set; } = string.Empty;
}
