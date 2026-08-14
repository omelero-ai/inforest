namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Opciones de configuración del PinPad.
/// Legacy: DLL3500.bas — IpPinPad, IpPort, nTimeOut, ruta de ConfigFileName.
/// </summary>
public sealed class PinPadOptions
{
    public const string SectionName = "Hardware:PinPad";

    /// <summary>Habilita el PinPad real. Si false, se usa NullPinPadService.</summary>
    public bool Habilitado { get; set; }

    /// <summary>Ruta al archivo de configuración de caja_pinpad.dll.</summary>
    public string ConfigPath { get; set; } = string.Empty;

    /// <summary>Ruta a la DLL nativa caja_pinpad.dll. Si vacío, se busca en el PATH del proceso.</summary>
    public string DllPath { get; set; } = string.Empty;
}
