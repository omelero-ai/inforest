using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Caja;

/// <summary>
/// Legacy: TMEDIODEPAGO, DPAGODOCUMENTO. BR-007, BR-013.
/// </summary>
public enum TipoMedioPago
{
    Efectivo = 1,
    Tarjeta = 2,
    Cheque = 3,
    Vale = 4,
    Otro = 5
}

/// <summary>
/// Maestro de medios de pago.
/// </summary>
public class MedioPago
{
    public string Codigo { get; private set; } = string.Empty;
    public string Descripcion { get; private set; } = string.Empty;
    public TipoMedioPago Tipo { get; private set; }
    public bool Activo { get; private set; }

    private MedioPago()
    {
    }

    public static MedioPago Crear(string codigo, string descripcion, TipoMedioPago tipo)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código del medio de pago es obligatorio.", "MEDIO_PAGO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(descripcion))
            throw new DomainException("La descripción del medio de pago es obligatoria.", "MEDIO_PAGO_DESCRIPCION_REQUERIDA");

        return new MedioPago
        {
            Codigo = codigo.Trim(),
            Descripcion = descripcion.Trim(),
            Tipo = tipo,
            Activo = true
        };
    }

    public static MedioPago Reconstituir(string codigo, string descripcion, TipoMedioPago tipo, bool activo)
    {
        return new MedioPago
        {
            Codigo = codigo,
            Descripcion = descripcion,
            Tipo = tipo,
            Activo = activo
        };
    }
}

/// <summary>
/// Pago registrado contra un documento.
/// </summary>
public class PagoDocumento
{
    public string CodigoDocumento { get; private set; } = string.Empty;
    public string CodigoMedioPago { get; private set; } = string.Empty;
    public decimal Monto { get; private set; }
    public decimal MontoCambio { get; private set; }
    public string? Referencia { get; private set; }
    public DateTime FechaRegistro { get; private set; }

    private PagoDocumento()
    {
    }

    public static PagoDocumento Crear(string documento, string medioPago, decimal monto, decimal cambio, string? referencia)
    {
        if (string.IsNullOrWhiteSpace(documento))
            throw new DomainException("El documento del pago es obligatorio.", "PAGO_DOCUMENTO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(medioPago))
            throw new DomainException("El medio de pago es obligatorio.", "PAGO_MEDIO_REQUERIDO");

        if (monto <= 0)
            throw new DomainException("El monto del pago debe ser mayor a cero.", "PAGO_MONTO_INVALIDO");

        if (cambio < 0)
            throw new DomainException("El cambio del pago no puede ser negativo.", "PAGO_CAMBIO_INVALIDO");

        return new PagoDocumento
        {
            CodigoDocumento = documento.Trim(),
            CodigoMedioPago = medioPago.Trim(),
            Monto = monto,
            MontoCambio = cambio,
            Referencia = string.IsNullOrWhiteSpace(referencia) ? null : referencia.Trim(),
            FechaRegistro = DateTime.Now
        };
    }

    public static PagoDocumento Reconstituir(
        string documento,
        string medioPago,
        decimal monto,
        decimal cambio,
        string? referencia,
        DateTime fechaRegistro)
    {
        return new PagoDocumento
        {
            CodigoDocumento = documento,
            CodigoMedioPago = medioPago,
            Monto = monto,
            MontoCambio = cambio,
            Referencia = referencia,
            FechaRegistro = fechaRegistro
        };
    }
}
