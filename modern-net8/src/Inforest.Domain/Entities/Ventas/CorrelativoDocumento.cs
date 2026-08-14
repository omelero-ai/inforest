using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Ventas;

/// <summary>
/// Registro de correlativo de comprobante por caja, impresora y tipo de emisión.
/// Legacy: TTIPODOCUMENTOIMPRESORA. BR-POS-008.
/// Tabla: INFOREST.dbo.TTIPODOCUMENTOIMPRESORA
/// </summary>
public class CorrelativoDocumento : Entity
{
    public string CodigoCaja { get; private set; } = string.Empty;
    public string CodigoImpresora { get; private set; } = string.Empty;
    public string TipoEmision { get; private set; } = string.Empty;
    public string? Descripcion { get; private set; }
    public string? Formulario { get; private set; }
    public string? Serie { get; private set; }
    public string? UltimoNumero { get; private set; }
    public string? NumeroAutorizacion { get; private set; }
    public DateTime? FechaInicio { get; private set; }
    public DateTime? FechaCaducidad { get; private set; }
    public bool Resumen { get; private set; }
    public bool Impuesto1 { get; private set; }
    public bool Impuesto2 { get; private set; }
    public bool Impuesto3 { get; private set; }
    public bool EquivaleDolares { get; private set; }
    public bool ImprimeImageCabecera { get; private set; }
    public bool ImprimeImagePie { get; private set; }
    public bool FacturacionElectronica { get; private set; }
    public string? UsuarioRegistro { get; private set; }
    public DateTime? FechaRegistro { get; private set; }

    private CorrelativoDocumento() { }

    public static CorrelativoDocumento Crear(
        string codigoCaja,
        string codigoImpresora,
        string tipoEmision,
        string? descripcion,
        string? formulario,
        string? serie,
        string? ultimoNumero,
        string? numeroAutorizacion,
        DateTime? fechaInicio,
        DateTime? fechaCaducidad,
        bool resumen,
        bool impuesto1,
        bool impuesto2,
        bool impuesto3,
        bool equivaleDolares,
        bool imprimeImageCabecera,
        bool imprimeImagePie,
        bool facturacionElectronica,
        string? usuarioRegistro,
        DateTime? fechaRegistro)
    {
        if (string.IsNullOrWhiteSpace(codigoCaja))
            throw new DomainException("El código de caja es requerido.", "CORRELATIVO_CAJA_REQUERIDA");
        if (string.IsNullOrWhiteSpace(codigoImpresora))
            throw new DomainException("El código de impresora es requerido.", "CORRELATIVO_IMPRESORA_REQUERIDA");
        if (string.IsNullOrWhiteSpace(tipoEmision))
            throw new DomainException("El tipo de emisión es requerido.", "CORRELATIVO_TIPO_EMISION_REQUERIDO");

        return new CorrelativoDocumento
        {
            CodigoCaja = codigoCaja.Trim(),
            CodigoImpresora = codigoImpresora.Trim(),
            TipoEmision = tipoEmision.Trim(),
            Descripcion = descripcion?.Trim(),
            Formulario = formulario?.Trim(),
            Serie = serie?.Trim(),
            UltimoNumero = ultimoNumero?.Trim(),
            NumeroAutorizacion = numeroAutorizacion?.Trim(),
            FechaInicio = fechaInicio,
            FechaCaducidad = fechaCaducidad,
            Resumen = resumen,
            Impuesto1 = impuesto1,
            Impuesto2 = impuesto2,
            Impuesto3 = impuesto3,
            EquivaleDolares = equivaleDolares,
            ImprimeImageCabecera = imprimeImageCabecera,
            ImprimeImagePie = imprimeImagePie,
            FacturacionElectronica = facturacionElectronica,
            UsuarioRegistro = usuarioRegistro?.Trim(),
            FechaRegistro = fechaRegistro
        };
    }

    /// <summary>
    /// Indica si este correlativo tiene un número válido asignado (BR-POS-008).
    /// Legacy: NFactura = "Sin Correlativo" si no existe registro en TTIPODOCUMENTOIMPRESORA.
    /// </summary>
    public bool TieneCorrelativo =>
        !string.IsNullOrWhiteSpace(UltimoNumero) &&
        !UltimoNumero.Equals("Sin Correlativo", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Indica si este correlativo está dentro del plazo de vigencia.
    /// Legacy: fCaducidad >= getdate()
    /// </summary>
    public bool EsVigente =>
        !FechaCaducidad.HasValue || FechaCaducidad.Value.Date >= DateTime.Today;

    /// <summary>
    /// Devuelve el siguiente número de comprobante como string.
    /// Legacy: tUltimoNumero incremental por caja.
    /// </summary>
    public string ObtenerNumeroActual() =>
        TieneCorrelativo ? UltimoNumero! : "Sin Correlativo";
}
