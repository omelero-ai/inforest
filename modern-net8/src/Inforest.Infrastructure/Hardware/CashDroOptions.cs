namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Opciones de configuración de CashDro.
/// Legacy: modProcedimientoNuevo.bas — App.path &amp; "\CashDrow\MotorCashDrow.exe".
/// </summary>
public sealed class CashDroOptions
{
    public const string SectionName = "Hardware:CashDro";

    /// <summary>Habilita el servicio CashDro real. Si false, se usa NullCashDroService.</summary>
    public bool Habilitado { get; set; }

    /// <summary>Ruta completa al ejecutable MotorCashDrow.exe.</summary>
    public string RutaMotor { get; set; } = string.Empty;

    /// <summary>Argumentos para la operación de apertura de cajón.</summary>
    public string ArgumentosAbrirCajon { get; set; } = "ABRIR";
}
