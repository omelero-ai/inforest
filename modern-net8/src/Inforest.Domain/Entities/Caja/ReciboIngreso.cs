using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Caja;

/// <summary>
/// Recibo de ingreso de caja (cabecera).
/// Legacy: MINGRESO. BR-RECIBO-001, BR-RECIBO-002, BR-RECIBO-003,
///         BR-RECIBO-004, BR-RECIBO-005, BR-RECIBO-006.
/// </summary>
public class ReciboIngreso : Entity
{
    /// <summary>Código del recibo: YY + 8 dígitos. BR-RECIBO-001.</summary>
    public string CodigoRecibo { get; private set; } = string.Empty;
    public DateTime Fecha { get; private set; }
    /// <summary>Moneda: '01'=MN, '02'=ME.</summary>
    public string Moneda { get; private set; } = "01";
    /// <summary>Tipo pago: '01'=Efectivo, '02'=Tarjeta, '03'=Cheque. BR-RECIBO-004.</summary>
    public string TipoPago { get; private set; } = "01";
    public string? CodigoTarjeta { get; private set; }
    public string? Referencia { get; private set; }
    public decimal TipoCambio { get; private set; }
    public decimal Monto { get; private set; }
    public string Descripcion { get; private set; } = string.Empty;
    public string? Autorizacion { get; private set; }
    /// <summary>Indica si el recibo corresponde a un anticipo. BR-RECIBO-005.</summary>
    public bool EsAnticipo { get; private set; }
    /// <summary>Estado: '01'=EMITIDO, '04'=ANULADO, '02'=PAGADO. BR-RECIBO-002.</summary>
    public string Estado { get; private set; } = "01";
    public string? CodigoTurno { get; private set; }
    public string? CodigoCaja { get; private set; }
    public string? Usuario { get; private set; }
    public DateTime FechaRegistro { get; private set; }
    public DateTime? FechaDiaContable { get; private set; }
    public string? CodigoPedido { get; private set; }

    private ReciboIngreso() { }

    /// <summary>
    /// Registra un nuevo recibo de ingreso.
    /// BR-RECIBO-001: Código obligatorio (formato YY+8 calculado en handler/repositorio).
    /// BR-RECIBO-011: Monto > 0 y descripción requeridos.
    /// </summary>
    public static ReciboIngreso Registrar(
        string codigo,
        DateTime fecha,
        decimal monto,
        string descripcion,
        string moneda = "01",
        string tipoPago = "01",
        string? codigoTarjeta = null,
        string? referencia = null,
        decimal tipoCambio = 1m,
        string? autorizacion = null,
        bool esAnticipo = false,
        string? turno = null,
        string? caja = null,
        string? usuario = null,
        DateTime? fechaDiaContable = null,
        string? codigoPedido = null)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código del recibo de ingreso es obligatorio.", "RECIBO_CODIGO_REQUERIDO");

        if (monto <= 0)
            throw new DomainException("El monto del recibo de ingreso debe ser mayor a cero.", "RECIBO_MONTO_INVALIDO");

        if (string.IsNullOrWhiteSpace(descripcion))
            throw new DomainException("La descripción del recibo de ingreso es obligatoria.", "RECIBO_DESCRIPCION_REQUERIDA");

        return new ReciboIngreso
        {
            CodigoRecibo = codigo,
            Fecha = fecha,
            Monto = monto,
            Descripcion = descripcion,
            Moneda = moneda,
            TipoPago = tipoPago,
            CodigoTarjeta = codigoTarjeta,
            Referencia = referencia,
            TipoCambio = tipoCambio,
            Autorizacion = autorizacion,
            EsAnticipo = esAnticipo,
            Estado = "01",
            CodigoTurno = turno,
            CodigoCaja = caja,
            Usuario = usuario,
            FechaRegistro = DateTime.Now,
            FechaDiaContable = fechaDiaContable,
            CodigoPedido = codigoPedido
        };
    }

    /// <summary>
    /// Reconstruye desde la base de datos (sin restricciones de estado).
    /// </summary>
    public static ReciboIngreso Reconstruir(
        string codigo,
        DateTime fecha,
        decimal monto,
        string descripcion,
        string moneda,
        string tipoPago,
        string? codigoTarjeta,
        string? referencia,
        decimal tipoCambio,
        string? autorizacion,
        bool esAnticipo,
        string estado,
        string? turno,
        string? caja,
        string? usuario,
        DateTime fechaRegistro,
        DateTime? fechaDiaContable,
        string? codigoPedido)
    {
        return new ReciboIngreso
        {
            CodigoRecibo = codigo,
            Fecha = fecha,
            Monto = monto,
            Descripcion = descripcion,
            Moneda = moneda,
            TipoPago = tipoPago,
            CodigoTarjeta = codigoTarjeta,
            Referencia = referencia,
            TipoCambio = tipoCambio,
            Autorizacion = autorizacion,
            EsAnticipo = esAnticipo,
            Estado = estado,
            CodigoTurno = turno,
            CodigoCaja = caja,
            Usuario = usuario,
            FechaRegistro = fechaRegistro,
            FechaDiaContable = fechaDiaContable,
            CodigoPedido = codigoPedido
        };
    }

    /// <summary>
    /// Anula el recibo. BR-RECIBO-003: No puede anularse si ya está anulado o pagado.
    /// </summary>
    public void Anular()
    {
        if (Estado == "04")
            throw new DomainException("El recibo de ingreso ya está anulado.", "RECIBO_YA_ANULADO");

        if (Estado == "02")
            throw new DomainException("No se puede anular un recibo de ingreso pagado.", "RECIBO_PAGADO_NO_ANULABLE");

        Estado = "04";
    }

    public bool EstaEmitido => Estado == "01";
    public bool EstaAnulado => Estado == "04";
    public bool EstaPagado => Estado == "02";

    /// <summary>
    /// Retorna la descripción legible del estado. BR-RECIBO-002.
    /// </summary>
    public string DescripcionEstado => Estado switch
    {
        "01" => "EMITIDO",
        "02" => "PAGADO",
        "04" => "ANULADO",
        _ => Estado
    };
}
