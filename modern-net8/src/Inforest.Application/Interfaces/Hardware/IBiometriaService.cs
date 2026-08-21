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

    /// <summary>
    /// Identifica al usuario en base de datos por huella dactilar (matching 1:N).
    /// Legacy: frmVerificacionHuellaSup.frm — escanea huella, compara contra TUSUARIO.tHuella
    /// y retorna tResumido del usuario identificado (variable global sVar1).
    /// Regla BR-PERIPH-004.
    /// </summary>
    /// <param name="modulo">Código de módulo activo ("01" POS, "02" Administración, "03" Consultas).</param>
    /// <param name="cancellationToken">Token de cancelación.</param>
    Task<BiometriaIdentificacionResult> IdentificarUsuarioAsync(string modulo, CancellationToken cancellationToken = default);
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

/// <summary>
/// Resultado de identificación biométrica 1:N.
/// Legacy: frmVerificacionHuellaSup.frm — wenterHuellaSup + sVar1 (tResumido del usuario identificado).
/// Regla BR-PERIPH-004.
/// </summary>
public sealed record BiometriaIdentificacionResult(
    bool EsExitoso,
    int CodigoError,
    string Mensaje,
    string? LoginIdentificado = null)
{
    public static BiometriaIdentificacionResult Identificado(string login)
        => new(true, 0, "OK", login);

    public static BiometriaIdentificacionResult Fallido(int codigoError, string mensaje)
        => new(false, codigoError, mensaje);
}
