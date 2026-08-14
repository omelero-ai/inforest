namespace Inforest.Application.Interfaces.Hardware;

/// <summary>
/// Abstracción para el cajón automático CashDro.
/// Legacy: modProcedimientoNuevo.bas — IniciarMotorCashDrow(Argumentos), ShellExecute a MotorCashDrow.exe.
/// Regla BR-PERIPH-002. ADR-011.
/// </summary>
public interface ICashDroService
{
    /// <summary>
    /// Inicia el motor CashDro con los argumentos especificados.
    /// Legacy: ShellExecute(1, "Open", App.path &amp; "\CashDrow\MotorCashDrow.exe", Argumentos, "", 1)
    /// </summary>
    Task IniciarMotorAsync(string argumentos, CancellationToken cancellationToken = default);

    /// <summary>
    /// Abre el cajón automático CashDro.
    /// </summary>
    Task AbrirCajonAsync(CancellationToken cancellationToken = default);
}
