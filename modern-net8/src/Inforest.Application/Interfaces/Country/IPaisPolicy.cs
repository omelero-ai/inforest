namespace Inforest.Application.Interfaces.Country;

/// <summary>
/// Política de país: expone flags y formatos fiscales/regulatorios por país.
/// Legacy: variable global `pais` en modDeclaracion.bas; scripts opcionales por país.
/// Regla BR-PERIPH-007. ADR-008, ADR-010.
/// </summary>
public interface IPaisPolicy
{
    /// <summary>Código de país (ej. "PE" = Perú, "AR" = Argentina, "CL" = Chile, "EC" = Ecuador, "BO" = Bolivia, "ES" = España).</summary>
    string CodigoPais { get; }

    /// <summary>Nombre del país.</summary>
    string Nombre { get; }

    /// <summary>Indica si el país requiere impresora fiscal (ej. Argentina con IFEpson.ocx).</summary>
    bool RequiereFiscal { get; }

    /// <summary>Indica si el país requiere Facturación Electrónica.</summary>
    bool RequiereFE { get; }

    /// <summary>
    /// Formato/etiqueta del número de documento tributario.
    /// Legacy: "R.U.C." (Perú), "C.U.I.T." (Argentina), "R.U.T." (Chile), "RUC" (Ecuador), "NIT" (Bolivia).
    /// </summary>
    string LabelDocumentoTributario { get; }

    /// <summary>Tipo/nombre del gateway de FE usado por este país (ej. "SUNAT", "AFIP", "SII", "SRI", "SIAT").</summary>
    string TipoFEGateway { get; }

    /// <summary>Código de moneda ISO 4217 (ej. "PEN", "ARS", "CLP", "USD", "BOB", "EUR").</summary>
    string CodigoMoneda { get; }

    /// <summary>Timezone fiscal del país (ej. "America/Lima", "America/Argentina/Buenos_Aires").</summary>
    string TimezoneFiscal { get; }
}
