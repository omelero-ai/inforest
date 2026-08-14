namespace Inforest.Application.Interfaces.Hardware;

/// <summary>
/// Abstracción para el servicio de biometría SecuGen.
/// Legacy: FpLibX_Const.bas, sgfplibx.ocx — captura y verificación de huella dactilar.
/// Regla BR-PERIPH-004. ADR-009.
/// </summary>
public interface IBiometriaService
{
    /// <summary>
    /// Captura la huella del lector biométrico conectado.
    /// Legacy: SDK FpLibX — captura y extracción de template.
    /// </summary>
    Task<BiometriaResult> CapturarHuellaAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifica si la huella capturada corresponde al usuario indicado.
    /// Legacy: SDK FpLibX — matching de template almacenado vs capturado.
    /// </summary>
    Task<BiometriaResult> VerificarUsuarioAsync(string codigoUsuario, byte[] templateHuella, CancellationToken cancellationToken = default);
}

/// <summary>
/// Resultado de una operación biométrica.
/// Legacy: ERROR_NONE=0, ERROR_EXTRACT_FAIL=105, ERROR_MATCH_FAIL=106 (FpLibX_Const.bas).
/// </summary>
public sealed record BiometriaResult(
    bool EsExitoso,
    int CodigoError,
    string Mensaje,
    int PuntuacionCoincidencia = 0,
    byte[]? Template = null)
{
    public static BiometriaResult Exitoso(int puntuacion, byte[]? template = null)
        => new(true, 0, "OK", puntuacion, template);

    public static BiometriaResult Fallido(int codigoError, string mensaje)
        => new(false, codigoError, mensaje);
}
