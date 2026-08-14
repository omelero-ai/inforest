using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Ventas;

/// <summary>
/// Nota de crédito emitida contra un documento de venta.
/// Legacy: MNOTACREDITO. BR-NC-001, BR-NC-002, BR-NC-003, BR-NC-004.
/// </summary>
public class NotaCredito : Entity
{
    public string CodigoNotaCredito { get; private set; } = string.Empty;
    public DateTime Fecha { get; private set; }
    public string CodigoDocumento { get; private set; } = string.Empty;
    public decimal Neto { get; private set; }
    public decimal Impuesto1 { get; private set; }
    public decimal Impuesto2 { get; private set; }
    public decimal Impuesto3 { get; private set; }
    public decimal Total { get; private set; }
    public string Estado { get; private set; } = string.Empty;
    public string? CodigoTurno { get; private set; }
    public string? CodigoCaja { get; private set; }
    public string? Usuario { get; private set; }
    public DateTime FechaRegistro { get; private set; }
    public string Observacion { get; private set; } = string.Empty;
    public string? TipoDocumento { get; private set; }
    public string? Motivo { get; private set; }
    public string? UsuarioAnulado { get; private set; }
    public DateTime? FechaAnulado { get; private set; }
    public DateTime? FechaDiaContable { get; private set; }

    private NotaCredito() { }

    /// <summary>
    /// Emite una nota de crédito contra un documento existente.
    /// BR-NC-001: Total > 0.
    /// BR-NC-002: Impuestos >= 0.
    /// BR-NC-004: Observacion requerida.
    /// </summary>
    public static NotaCredito Emitir(
        string codigo,
        DateTime fecha,
        string codigoDocumento,
        decimal neto,
        decimal impuesto1,
        decimal impuesto2,
        decimal impuesto3,
        string observacion,
        string? turno = null,
        string? caja = null,
        string? usuario = null,
        string? tipoDocumento = null,
        string? motivo = null)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código de la nota de crédito es obligatorio.", "NC_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoDocumento))
            throw new DomainException("El código del documento afectado es obligatorio.", "NC_DOCUMENTO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(observacion))
            throw new DomainException("El motivo de la nota de crédito es obligatorio.", "NC_OBSERVACION_REQUERIDA");

        // BR-NC-002: impuestos no negativos
        if (impuesto1 < 0 || impuesto2 < 0 || impuesto3 < 0)
            throw new DomainException("Los impuestos de la nota de crédito no pueden ser negativos.", "NC_IMPUESTO_NEGATIVO");

        var total = neto + impuesto1 + impuesto2 + impuesto3;

        // BR-NC-001: total > 0
        if (total <= 0)
            throw new DomainException("El valor de la nota de crédito debe ser mayor a cero.", "NC_TOTAL_INVALIDO");

        return new NotaCredito
        {
            CodigoNotaCredito = codigo.Trim(),
            Fecha = fecha,
            CodigoDocumento = codigoDocumento.Trim(),
            Neto = neto,
            Impuesto1 = impuesto1,
            Impuesto2 = impuesto2,
            Impuesto3 = impuesto3,
            Total = total,
            Estado = "PE",
            CodigoTurno = turno?.Trim(),
            CodigoCaja = caja?.Trim(),
            Usuario = usuario?.Trim(),
            FechaRegistro = DateTime.Now,
            FechaDiaContable = DateTime.Today,
            Observacion = observacion.Trim(),
            TipoDocumento = tipoDocumento?.Trim(),
            Motivo = motivo?.Trim()
        };
    }

    /// <summary>
    /// Reconstituye una nota de crédito desde persistencia.
    /// </summary>
    public static NotaCredito Reconstituir(
        string codigo,
        DateTime fecha,
        string codigoDocumento,
        decimal neto,
        decimal impuesto1,
        decimal impuesto2,
        decimal impuesto3,
        decimal total,
        string estado,
        string? turno,
        string? caja,
        string? usuario,
        DateTime fechaRegistro,
        DateTime? fechaDiaContable,
        string observacion,
        string? tipoDocumento,
        string? motivo,
        string? usuarioAnulado,
        DateTime? fechaAnulado)
    {
        return new NotaCredito
        {
            CodigoNotaCredito = codigo,
            Fecha = fecha,
            CodigoDocumento = codigoDocumento,
            Neto = neto,
            Impuesto1 = impuesto1,
            Impuesto2 = impuesto2,
            Impuesto3 = impuesto3,
            Total = total,
            Estado = estado,
            CodigoTurno = turno,
            CodigoCaja = caja,
            Usuario = usuario,
            FechaRegistro = fechaRegistro,
            FechaDiaContable = fechaDiaContable,
            Observacion = observacion,
            TipoDocumento = tipoDocumento,
            Motivo = motivo,
            UsuarioAnulado = usuarioAnulado,
            FechaAnulado = fechaAnulado
        };
    }

    /// <summary>
    /// Anula la nota de crédito.
    /// BR-NC-006: No se puede anular una NC ya anulada.
    /// </summary>
    public void Anular(string usuario)
    {
        if (Estado == "AN")
            throw new DomainException("La nota de crédito ya se encuentra anulada.", "NC_YA_ANULADA");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario de anulación es obligatorio.", "NC_USUARIO_ANULACION_REQUERIDO");

        Estado = "AN";
        UsuarioAnulado = usuario.Trim();
        FechaAnulado = DateTime.Now;
    }
}
