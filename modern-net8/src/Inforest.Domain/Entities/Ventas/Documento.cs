using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Ventas;

/// <summary>
/// Documento de venta.
/// Legacy: MDOCUMENTO. BR-002, BR-013.
/// </summary>
public class Documento : Entity
{
    public string CodigoDocumento { get; private set; } = string.Empty;
    public string TipoDocumento { get; private set; } = string.Empty;
    public string? CodigoCliente { get; private set; }
    public decimal Neto { get; private set; }
    public decimal Impuesto1 { get; private set; }
    public decimal Impuesto2 { get; private set; }
    public decimal Impuesto3 { get; private set; }
    public decimal Venta { get; private set; }
    public decimal Descuento { get; private set; }
    public decimal Propina { get; private set; }
    public decimal Total { get; private set; }
    public decimal Abono { get; private set; }
    public decimal Vuelto { get; private set; }
    public string Estado { get; private set; } = string.Empty;
    public string? CodigoTurno { get; private set; }
    public string? CodigoCaja { get; private set; }
    public DateTime? FechaPago { get; private set; }
    public DateTime FechaRegistro { get; private set; }
    public DateTime? FechaDiaContable { get; private set; }
    public string? UsuarioAnulado { get; private set; }
    public string? MotivoAnulacion { get; private set; }
    public DateTime? FechaAnulacion { get; private set; }

    private Documento()
    {
    }

    public static Documento Emitir(
        string codigo,
        string tipo,
        string? cliente,
        decimal neto,
        decimal impuesto1,
        decimal impuesto2,
        decimal impuesto3,
        decimal propina,
        decimal descuento,
        string turno,
        string caja)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código del documento es obligatorio.", "DOCUMENTO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(tipo))
            throw new DomainException("El tipo de documento es obligatorio.", "DOCUMENTO_TIPO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(turno))
            throw new DomainException("El turno es obligatorio para emitir un documento.", "DOCUMENTO_TURNO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(caja))
            throw new DomainException("La caja es obligatoria para emitir un documento.", "DOCUMENTO_CAJA_REQUERIDA");

        if (neto < 0 || impuesto1 < 0 || impuesto2 < 0 || impuesto3 < 0 || propina < 0 || descuento < 0)
            throw new DomainException("Los importes del documento no pueden ser negativos.", "DOCUMENTO_MONTO_INVALIDO");

        var venta = neto + impuesto1 + impuesto2 + impuesto3;
        var documento = new Documento
        {
            CodigoDocumento = codigo.Trim(),
            TipoDocumento = tipo.Trim(),
            CodigoCliente = string.IsNullOrWhiteSpace(cliente) ? null : cliente.Trim(),
            Neto = neto,
            Impuesto1 = impuesto1,
            Impuesto2 = impuesto2,
            Impuesto3 = impuesto3,
            Venta = venta,
            Descuento = descuento,
            Propina = propina,
            Estado = "PE",
            CodigoTurno = turno.Trim(),
            CodigoCaja = caja.Trim(),
            FechaRegistro = DateTime.Now,
            FechaDiaContable = DateTime.Today
        };

        documento.Total = documento.CalcularTotal();
        return documento;
    }

    public static Documento Reconstituir(
        string codigo,
        string tipo,
        string? cliente,
        decimal neto,
        decimal impuesto1,
        decimal impuesto2,
        decimal impuesto3,
        decimal venta,
        decimal descuento,
        decimal propina,
        decimal total,
        string estado,
        string? turno,
        string? caja,
        DateTime fechaRegistro,
        DateTime? fechaPago,
        DateTime? fechaDiaContable,
        decimal abono = 0,
        decimal vuelto = 0,
        string? usuarioAnulado = null,
        string? motivoAnulacion = null,
        DateTime? fechaAnulacion = null)
    {
        return new Documento
        {
            CodigoDocumento = codigo,
            TipoDocumento = tipo,
            CodigoCliente = cliente,
            Neto = neto,
            Impuesto1 = impuesto1,
            Impuesto2 = impuesto2,
            Impuesto3 = impuesto3,
            Venta = venta,
            Descuento = descuento,
            Propina = propina,
            Total = total,
            Estado = estado,
            CodigoTurno = turno,
            CodigoCaja = caja,
            FechaRegistro = fechaRegistro,
            FechaPago = fechaPago,
            FechaDiaContable = fechaDiaContable,
            Abono = abono,
            Vuelto = vuelto,
            UsuarioAnulado = usuarioAnulado,
            MotivoAnulacion = motivoAnulacion,
            FechaAnulacion = fechaAnulacion
        };
    }

    public void Cobrar(decimal abono)
    {
        if (Estado == "CO")
            throw new DomainException("El documento ya fue cobrado.", "DOCUMENTO_YA_COBRADO");

        if (Estado == "AN")
            throw new DomainException("No se puede cobrar un documento anulado.", "DOCUMENTO_ANULADO");

        if (abono <= 0)
            throw new DomainException("El abono debe ser mayor a cero.", "DOCUMENTO_ABONO_INVALIDO");

        if (abono < Total)
            throw new DomainException("El abono no cubre el total del documento.", "DOCUMENTO_ABONO_INSUFICIENTE");

        Abono = abono;
        Vuelto = abono - Total;
        FechaPago = DateTime.Now;
        Estado = "CO";
    }

    public void Anular(string usuario, string motivo)
    {
        if (Estado == "AN")
            throw new DomainException("El documento ya se encuentra anulado.", "DOCUMENTO_YA_ANULADO");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario de anulación es obligatorio.", "DOCUMENTO_USUARIO_ANULACION_REQUERIDO");

        if (string.IsNullOrWhiteSpace(motivo))
            throw new DomainException("El motivo de anulación es obligatorio.", "DOCUMENTO_MOTIVO_ANULACION_REQUERIDO");

        Estado = "AN";
        UsuarioAnulado = usuario.Trim();
        MotivoAnulacion = motivo.Trim();
        FechaAnulacion = DateTime.Now;
    }

    public decimal CalcularTotal() => Neto + Impuesto1 + Impuesto2 + Impuesto3 + Propina - Descuento;
}
