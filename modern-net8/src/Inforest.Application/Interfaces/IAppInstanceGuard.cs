namespace Inforest.Application.Interfaces;

/// <summary>
/// Legacy: modPuntoVenta.bas valida App.PrevInstance antes de iniciar el POS (BR-POS-001).
/// </summary>
public interface IAppInstanceGuard
{
    /// <summary>
    /// Intenta obtener el lock de instancia única para un identificador de aplicación.
    /// </summary>
    bool TryAcquire(string instanceKey, out IDisposable lease);
}
