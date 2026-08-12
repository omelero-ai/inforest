namespace Inforest.Domain.Entities.Hardware;

/// <summary>
/// Configuración de un país para periféricos y fiscalidad.
/// Legacy: variable global `pais` en modDeclaracion.bas; scripts opcionales por país.
/// Regla BR-PERIPH-007.
/// </summary>
public sealed class ConfiguracionPais
{
    /// <summary>Código de país ISO/Legacy (ej. "PE", "AR", "CL", "EC", "BO", "ES").</summary>
    public string CodigoPais { get; }

    /// <summary>Nombre del país.</summary>
    public string Nombre { get; }

    /// <summary>Etiqueta del documento tributario principal (RUC, CUIT, RUT, NIT).</summary>
    public string LabelDocumentoTributario { get; }

    /// <summary>Tipo de gateway de FE (SUNAT, AFIP, SII, SRI, SIAT, NONE).</summary>
    public string TipoFEGateway { get; }

    /// <summary>Código de moneda ISO 4217.</summary>
    public string CodigoMoneda { get; }

    /// <summary>Timezone fiscal.</summary>
    public string TimezoneFiscal { get; }

    /// <summary>Indica si este país requiere impresora fiscal.</summary>
    public bool RequiereFiscal { get; }

    /// <summary>Indica si este país requiere Facturación Electrónica.</summary>
    public bool RequiereFE { get; }

    private ConfiguracionPais(
        string codigoPais,
        string nombre,
        string labelDocumentoTributario,
        string tipoFEGateway,
        string codigoMoneda,
        string timezoneFiscal,
        bool requiereFiscal,
        bool requiereFE)
    {
        CodigoPais = codigoPais;
        Nombre = nombre;
        LabelDocumentoTributario = labelDocumentoTributario;
        TipoFEGateway = tipoFEGateway;
        CodigoMoneda = codigoMoneda;
        TimezoneFiscal = timezoneFiscal;
        RequiereFiscal = requiereFiscal;
        RequiereFE = requiereFE;
    }

    public static ConfiguracionPais Crear(
        string codigoPais,
        string nombre,
        string labelDocumentoTributario,
        string tipoFEGateway,
        string codigoMoneda,
        string timezoneFiscal,
        bool requiereFiscal,
        bool requiereFE)
    {
        if (string.IsNullOrWhiteSpace(codigoPais))
            throw new ArgumentException("El código de país es requerido.", nameof(codigoPais));

        return new ConfiguracionPais(
            codigoPais.ToUpperInvariant(),
            nombre,
            labelDocumentoTributario,
            tipoFEGateway,
            codigoMoneda,
            timezoneFiscal,
            requiereFiscal,
            requiereFE);
    }

    // Instancias predefinidas para los países soportados por el Legacy
    // Legacy: modDeclaracion.bas — `pais = "003"` para Argentina, default Perú

    /// <summary>Perú — RUC, SUNAT/OSE, PEN. Legacy: país por defecto.</summary>
    public static readonly ConfiguracionPais Peru = Crear(
        "PE", "Perú", "R.U.C.", "SUNAT", "PEN", "America/Lima", false, true);

    /// <summary>Argentina — CUIT, AFIP, ARS, con impresora fiscal Epson. Legacy: pais="003".</summary>
    public static readonly ConfiguracionPais Argentina = Crear(
        "AR", "Argentina", "C.U.I.T.", "AFIP", "ARS", "America/Argentina/Buenos_Aires", true, true);

    /// <summary>Chile — RUT, SII, CLP.</summary>
    public static readonly ConfiguracionPais Chile = Crear(
        "CL", "Chile", "R.U.T.", "SII", "CLP", "America/Santiago", false, true);

    /// <summary>Ecuador — RUC, SRI, USD.</summary>
    public static readonly ConfiguracionPais Ecuador = Crear(
        "EC", "Ecuador", "R.U.C.", "SRI", "USD", "America/Guayaquil", false, true);

    /// <summary>Bolivia — NIT, SIAT, BOB.</summary>
    public static readonly ConfiguracionPais Bolivia = Crear(
        "BO", "Bolivia", "N.I.T.", "SIAT", "BOB", "America/La_Paz", false, true);

    /// <summary>España — NIF, AEAT, EUR.</summary>
    public static readonly ConfiguracionPais Espana = Crear(
        "ES", "España", "N.I.F.", "AEAT", "EUR", "Europe/Madrid", false, false);

    /// <summary>Devuelve la configuración predefinida por código, o Perú si no se reconoce.</summary>
    public static ConfiguracionPais PorCodigo(string codigoPais) => codigoPais.ToUpperInvariant() switch
    {
        "PE" => Peru,
        "AR" => Argentina,
        "CL" => Chile,
        "EC" => Ecuador,
        "BO" => Bolivia,
        "ES" => Espana,
        _ => Peru
    };
}
