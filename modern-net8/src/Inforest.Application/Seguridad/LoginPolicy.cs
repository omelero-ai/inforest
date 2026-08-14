namespace Inforest.Application.Seguridad;

/// <summary>
/// Política de login para el POS.
/// Legacy: frmAcceso.frm — reglas de acceso y bloqueo.
/// </summary>
public static class LoginPolicy
{
    /// <summary>
    /// Número máximo de intentos fallidos antes de terminar la aplicación.
    /// Legacy: frmAcceso.frm — contador <c>i</c> inicia en 1 y se incrementa por fallo;
    /// cuando <c>i = 4</c> se ejecuta <c>End</c>. Equivale a 3 intentos fallidos.
    /// BR-POS-006-LOCK.
    /// </summary>
    public const int MaxIntentosFallidos = 3;
}
