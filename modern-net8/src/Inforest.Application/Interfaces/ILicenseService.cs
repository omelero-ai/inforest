namespace Inforest.Application.Interfaces;

/// <summary>
/// Servicio de validación de licencia/transición desde THARDKEY.
/// Legacy: License.cls + modSeguridadInfhotel.bas.
/// </summary>
public interface ILicenseService
{
    Task<LicenseValidationResult> ValidarLicenciaAsync(LicenseValidationRequest request, CancellationToken cancellationToken = default);
}

public sealed record LicenseValidationRequest(
    bool Habilitado,
    string ModuloEjecutable,
    string BaseDatos,
    string? RucEmpresa,
    bool OmitirHardwareLegacy = true);

public sealed record LicenseValidationResult(
    bool EsValida,
    string? CodigoError,
    string? Mensaje,
    string? Advertencia,
    int? DiasRestantes,
    string Estrategia)
{
    public static LicenseValidationResult Valida(string estrategia, string? advertencia = null, int? diasRestantes = null)
        => new(true, null, null, advertencia, diasRestantes, estrategia);

    public static LicenseValidationResult Invalida(string codigoError, string mensaje, string estrategia)
        => new(false, codigoError, mensaje, null, null, estrategia);
}
