namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Opciones de configuración de la Impresora Fiscal Epson (Argentina).
/// Legacy: modImpresoraFiscal.bas — IFEpson.ocx, configuración AFIP Argentina.
/// </summary>
public sealed class ImpresoraFiscalOptions
{
    public const string SectionName = "Hardware:ImpresoraFiscal";

    /// <summary>Habilita la impresora fiscal. Si false, se usa NullImpresoraFiscalService.</summary>
    public bool Habilitado { get; set; }

    /// <summary>Puerto serial de la impresora fiscal (ej. "COM1", "COM3").</summary>
    public string Puerto { get; set; } = "COM1";
}

/// <summary>
/// Opciones de configuración de biometría SecuGen.
/// Legacy: FpLibX_Const.bas — sgfplibx.ocx.
/// </summary>
public sealed class BiometriaOptions
{
    public const string SectionName = "Hardware:Biometria";

    /// <summary>Habilita el lector biométrico. Si false, se usa NullBiometriaService.</summary>
    public bool Habilitado { get; set; }
}
