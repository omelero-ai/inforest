using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Caja;

/// <summary>
/// Recibo de egreso de caja.
/// Legacy: MEGRESO. BR-RECIBO-007, BR-RECIBO-008, BR-RECIBO-009,
///         BR-RECIBO-010, BR-RECIBO-011, BR-RECIBO-012.
/// </summary>
public class ReciboEgreso : Entity
{
    /// <summary>Código del recibo: YY + 8 dígitos. BR-RECIBO-007.</summary>
    public string CodigoRecibo { get; private set; } = string.Empty;
    public string? CodigoCaja { get; private set; }
    public string? CodigoTurno { get; private set; }
    public DateTime Fecha { get; private set; }
    /// <summary>Moneda: '01'=MN, '02'=ME.</summary>
    public string Moneda { get; private set; } = "01";
    public decimal TipoCambio { get; private set; }
    public decimal Monto { get; private set; }
    public string Descripcion { get; private set; } = string.Empty;
    public string? Autorizacion { get; private set; }
    /// <summary>Estado: '01'=EMITIDO, '04'=ANULADO. BR-RECIBO-008.</summary>
    public string Estado { get; private set; } = "01";
    public string? Usuario { get; private set; }
    public DateTime FechaRegistro { get; private set; }
    public DateTime? FechaDiaContable { get; private set; }
    public string? TipoEgreso { get; private set; }
    /// <summary>Referencia a nota de crédito o comprobante externo.</summary>
    public string? Referencia { get; private set; }
    /// <summary>Modo de egreso: VARIOS, NOTACREDITO, MERCADERIA. BR-RECIBO-010.</summary>
    public string ModoEgreso { get; private set; } = "VARIOS";
    public string? TipoIdentificacion { get; private set; }
    public string? Identificacion { get; private set; }
    public string? RazonSocial { get; private set; }
    public string? NumeroComprobante { get; private set; }
    public string? TipoComprobante { get; private set; }
    public string? CentroCosto { get; private set; }
    public string? Observacion { get; private set; }

    private ReciboEgreso() { }

    /// <summary>
    /// Registra un nuevo recibo de egreso.
    /// BR-RECIBO-011: Monto > 0 y descripción requeridos.
    /// BR-RECIBO-012: Para NOTACREDITO la referencia es obligatoria.
    /// </summary>
    public static ReciboEgreso Registrar(
        string codigo,
        DateTime fecha,
        decimal monto,
        string descripcion,
        string modoEgreso = "VARIOS",
        string moneda = "01",
        decimal tipoCambio = 1m,
        string? autorizacion = null,
        string? tipoEgreso = null,
        string? referencia = null,
        string? turno = null,
        string? caja = null,
        string? usuario = null,
        DateTime? fechaDiaContable = null,
        string? tipoIdentificacion = null,
        string? identificacion = null,
        string? razonSocial = null,
        string? numeroComprobante = null,
        string? tipoComprobante = null,
        string? centroCosto = null,
        string? observacion = null)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código del recibo de egreso es obligatorio.", "EGRESO_CODIGO_REQUERIDO");

        if (monto <= 0)
            throw new DomainException("El monto del recibo de egreso debe ser mayor a cero.", "EGRESO_MONTO_INVALIDO");

        if (string.IsNullOrWhiteSpace(descripcion))
            throw new DomainException("La descripción del recibo de egreso es obligatoria.", "EGRESO_DESCRIPCION_REQUERIDA");

        // BR-RECIBO-012: ModoEgreso NOTACREDITO requiere referencia
        if (modoEgreso == "NOTACREDITO" && string.IsNullOrWhiteSpace(referencia))
            throw new DomainException("Para un egreso tipo Nota de Crédito se requiere la referencia.", "EGRESO_REFERENCIA_NC_REQUERIDA");

        // BR-RECIBO-010: ModoEgreso válido
        if (modoEgreso != "VARIOS" && modoEgreso != "NOTACREDITO" && modoEgreso != "MERCADERIA")
            throw new DomainException($"Modo de egreso '{modoEgreso}' no reconocido.", "EGRESO_MODO_INVALIDO");

        return new ReciboEgreso
        {
            CodigoRecibo = codigo,
            Fecha = fecha,
            Monto = monto,
            Descripcion = descripcion,
            ModoEgreso = modoEgreso,
            Moneda = moneda,
            TipoCambio = tipoCambio,
            Autorizacion = autorizacion,
            TipoEgreso = tipoEgreso,
            Referencia = referencia,
            Estado = "01",
            CodigoTurno = turno,
            CodigoCaja = caja,
            Usuario = usuario,
            FechaRegistro = DateTime.Now,
            FechaDiaContable = fechaDiaContable,
            TipoIdentificacion = tipoIdentificacion,
            Identificacion = identificacion,
            RazonSocial = razonSocial,
            NumeroComprobante = numeroComprobante,
            TipoComprobante = tipoComprobante,
            CentroCosto = centroCosto,
            Observacion = observacion
        };
    }

    /// <summary>
    /// Reconstruye desde la base de datos.
    /// </summary>
    public static ReciboEgreso Reconstruir(
        string codigo,
        string? caja,
        string? turno,
        DateTime fecha,
        string moneda,
        decimal tipoCambio,
        decimal monto,
        string descripcion,
        string? autorizacion,
        string estado,
        string? usuario,
        DateTime fechaRegistro,
        DateTime? fechaDiaContable,
        string? tipoEgreso,
        string? referencia,
        string modoEgreso = "VARIOS",
        string? tipoIdentificacion = null,
        string? identificacion = null,
        string? razonSocial = null,
        string? numeroComprobante = null,
        string? tipoComprobante = null,
        string? centroCosto = null,
        string? observacion = null)
    {
        return new ReciboEgreso
        {
            CodigoRecibo = codigo,
            CodigoCaja = caja,
            CodigoTurno = turno,
            Fecha = fecha,
            Moneda = moneda,
            TipoCambio = tipoCambio,
            Monto = monto,
            Descripcion = descripcion,
            Autorizacion = autorizacion,
            Estado = estado,
            Usuario = usuario,
            FechaRegistro = fechaRegistro,
            FechaDiaContable = fechaDiaContable,
            TipoEgreso = tipoEgreso,
            Referencia = referencia,
            ModoEgreso = modoEgreso,
            TipoIdentificacion = tipoIdentificacion,
            Identificacion = identificacion,
            RazonSocial = razonSocial,
            NumeroComprobante = numeroComprobante,
            TipoComprobante = tipoComprobante,
            CentroCosto = centroCosto,
            Observacion = observacion
        };
    }

    /// <summary>
    /// Anula el egreso. BR-RECIBO-008: No puede anularse si ya está anulado.
    /// </summary>
    public void Anular()
    {
        if (Estado == "04")
            throw new DomainException("El recibo de egreso ya está anulado.", "EGRESO_YA_ANULADO");

        Estado = "04";
    }

    public bool EstaEmitido => Estado == "01";
    public bool EstaAnulado => Estado == "04";

    /// <summary>
    /// Retorna la descripción legible del estado. BR-RECIBO-008.
    /// </summary>
    public string DescripcionEstado => Estado switch
    {
        "01" => "EMITIDO",
        "04" => "ANULADO",
        _ => Estado
    };
}
