using Inforest.Application.Interfaces.Country;

namespace Inforest.Infrastructure.Country;

/// <summary>
/// Implementación de política de país para Perú.
/// Legacy: país por defecto (pais = "" / "PE"). R.U.C., SUNAT, PEN.
/// Regla BR-PERIPH-007. ADR-008.
/// </summary>
internal sealed class PeruPaisPolicy : IPaisPolicy
{
    public string CodigoPais => "PE";
    public string Nombre => "Perú";
    public bool RequiereFiscal => false;
    public bool RequiereFE => true;
    public string LabelDocumentoTributario => "R.U.C.";
    public string TipoFEGateway => "SUNAT";
    public string CodigoMoneda => "PEN";
    public string TimezoneFiscal => "America/Lima";
}

/// <summary>
/// Implementación de política de país para Argentina.
/// Legacy: pais = "003". C.U.I.T., AFIP, ARS. Impresora Fiscal Epson requerida.
/// Regla BR-PERIPH-007. ADR-008.
/// </summary>
internal sealed class ArgentinaPaisPolicy : IPaisPolicy
{
    public string CodigoPais => "AR";
    public string Nombre => "Argentina";
    public bool RequiereFiscal => true;
    public bool RequiereFE => true;
    public string LabelDocumentoTributario => "C.U.I.T.";
    public string TipoFEGateway => "AFIP";
    public string CodigoMoneda => "ARS";
    public string TimezoneFiscal => "America/Argentina/Buenos_Aires";
}

/// <summary>
/// Implementación de política de país para Chile.
/// Legacy: scriptChileAlIniciar.sql. R.U.T., SII, CLP.
/// Regla BR-PERIPH-007. ADR-008.
/// </summary>
internal sealed class ChilePaisPolicy : IPaisPolicy
{
    public string CodigoPais => "CL";
    public string Nombre => "Chile";
    public bool RequiereFiscal => false;
    public bool RequiereFE => true;
    public string LabelDocumentoTributario => "R.U.T.";
    public string TipoFEGateway => "SII";
    public string CodigoMoneda => "CLP";
    public string TimezoneFiscal => "America/Santiago";
}

/// <summary>
/// Implementación de política de país para Ecuador.
/// Legacy: scriptEcuadorAlIniciar.sql. R.U.C., SRI, USD.
/// Regla BR-PERIPH-007. ADR-008.
/// </summary>
internal sealed class EcuadorPaisPolicy : IPaisPolicy
{
    public string CodigoPais => "EC";
    public string Nombre => "Ecuador";
    public bool RequiereFiscal => false;
    public bool RequiereFE => true;
    public string LabelDocumentoTributario => "R.U.C.";
    public string TipoFEGateway => "SRI";
    public string CodigoMoneda => "USD";
    public string TimezoneFiscal => "America/Guayaquil";
}

/// <summary>
/// Implementación de política de país para Bolivia.
/// Legacy: scriptBoliviaAlIniciar.sql. N.I.T., SIAT, BOB.
/// Regla BR-PERIPH-007. ADR-008.
/// </summary>
internal sealed class BoliviaPaisPolicy : IPaisPolicy
{
    public string CodigoPais => "BO";
    public string Nombre => "Bolivia";
    public bool RequiereFiscal => false;
    public bool RequiereFE => true;
    public string LabelDocumentoTributario => "N.I.T.";
    public string TipoFEGateway => "SIAT";
    public string CodigoMoneda => "BOB";
    public string TimezoneFiscal => "America/La_Paz";
}

/// <summary>
/// Implementación de política de país para España.
/// Legacy: scriptEspanaAlIniciar.sql. N.I.F., AEAT, EUR.
/// Regla BR-PERIPH-007. ADR-008.
/// </summary>
internal sealed class EspanaPaisPolicy : IPaisPolicy
{
    public string CodigoPais => "ES";
    public string Nombre => "España";
    public bool RequiereFiscal => false;
    public bool RequiereFE => false;
    public string LabelDocumentoTributario => "N.I.F.";
    public string TipoFEGateway => "AEAT";
    public string CodigoMoneda => "EUR";
    public string TimezoneFiscal => "Europe/Madrid";
}

/// <summary>
/// Fábrica de políticas de país.
/// Resuelve la política correcta por código de país, con Perú como fallback.
/// Legacy: variable global `pais` en modDeclaracion.bas.
/// </summary>
public sealed class PaisPolicyFactory
{
    private readonly IReadOnlyDictionary<string, IPaisPolicy> _policies;

    public PaisPolicyFactory(IEnumerable<IPaisPolicy> policies)
    {
        _policies = policies.ToDictionary(p => p.CodigoPais, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Devuelve la política para el país indicado. Retorna Perú si el código no está registrado.</summary>
    public IPaisPolicy ObtenerPolitica(string codigoPais)
    {
        if (_policies.TryGetValue(codigoPais, out var policy))
            return policy;

        // Legacy: comportamiento por defecto = Perú
        return _policies.TryGetValue("PE", out var peru) ? peru
            : throw new InvalidOperationException("PeruPaisPolicy no está registrada.");
    }
}
