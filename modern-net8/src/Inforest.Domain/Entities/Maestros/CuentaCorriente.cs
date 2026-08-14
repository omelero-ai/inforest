using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Cuenta corriente habilitada para un cliente del restaurante.
/// Legacy: TDELIVERY (lClienteCtaCte=1) / vCompania / frmCtaCte.frm
/// BR-CTACTE-001, BR-CTACTE-002, BR-CTACTE-003
/// </summary>
public class CuentaCorriente : Entity
{
    /// <summary>Código TDELIVERY — clave de cuenta corriente.</summary>
    public string CodigoCtaCte { get; private set; } = string.Empty;

    /// <summary>Nombre o razón social (tApellido + ' ' + tNombre de TDELIVERY).</summary>
    public string Descripcion { get; private set; } = string.Empty;

    /// <summary>Identificación tributaria del cliente ctacte.</summary>
    public string? Identidad { get; private set; }

    /// <summary>Límite de crédito asignado (TDELIVERY.nLinea).</summary>
    public decimal LineaCredito { get; private set; }

    /// <summary>Consumo acumulado en cuenta corriente (TDELIVERY.nConsumo).</summary>
    public decimal Consumo { get; private set; }

    /// <summary>Saldo disponible = LineaCredito - Consumo (calculado en vCompania).</summary>
    public decimal Saldo => LineaCredito - Consumo;

    /// <summary>Días de crédito habilitados (TDELIVERY.nDiasCredito).</summary>
    public int DiasCredito { get; private set; }

    /// <summary>Tipo de cuenta corriente (TDELIVERY.tTipoCtaCte).</summary>
    public string? TipoCtaCte { get; private set; }

    /// <summary>Código TCLIENTE vinculado, si aplica.</summary>
    public string? CodigoCliente { get; private set; }

    public bool Activo { get; private set; }
    public string Usuario { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }

    private CuentaCorriente() { }

    public static CuentaCorriente Crear(
        string codigoCtaCte,
        string descripcion,
        decimal lineaCredito,
        decimal consumo,
        string usuario,
        string? identidad = null,
        int diasCredito = 0,
        string? tipoCtaCte = null,
        string? codigoCliente = null,
        bool activo = true,
        DateTime? fechaRegistro = null)
    {
        if (string.IsNullOrWhiteSpace(codigoCtaCte))
            throw new DomainException("El código de cuenta corriente es obligatorio.", "CTACTE_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(descripcion))
            throw new DomainException("La descripción de la cuenta corriente es obligatoria.", "CTACTE_DESCRIPCION_REQUERIDA");

        if (lineaCredito < 0)
            throw new DomainException("La línea de crédito no puede ser negativa.", "CTACTE_LINEA_INVALIDA");

        if (consumo < 0)
            throw new DomainException("El consumo no puede ser negativo.", "CTACTE_CONSUMO_INVALIDO");

        return new CuentaCorriente
        {
            CodigoCtaCte = codigoCtaCte.Trim(),
            Descripcion = descripcion.Trim(),
            LineaCredito = lineaCredito,
            Consumo = consumo,
            Usuario = usuario.Trim(),
            Identidad = string.IsNullOrWhiteSpace(identidad) ? null : identidad.Trim(),
            DiasCredito = diasCredito,
            TipoCtaCte = string.IsNullOrWhiteSpace(tipoCtaCte) ? null : tipoCtaCte.Trim(),
            CodigoCliente = string.IsNullOrWhiteSpace(codigoCliente) ? null : codigoCliente.Trim(),
            Activo = activo,
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };
    }

    /// <summary>
    /// Verifica si el consumo solicitado cabe dentro de la línea de crédito.
    /// BR-CTACTE-001: El consumo no puede superar la línea de crédito asignada.
    /// </summary>
    public bool TieneSaldoDisponible(decimal montoAdicional)
        => Consumo + montoAdicional <= LineaCredito;

    /// <summary>
    /// Aplica un consumo adicional a la cuenta.
    /// BR-CTACTE-001.
    /// </summary>
    public void AplicarConsumo(decimal monto)
    {
        if (monto <= 0)
            throw new DomainException("El monto a consumir debe ser mayor a cero.", "CTACTE_MONTO_INVALIDO");

        if (!TieneSaldoDisponible(monto))
            throw new DomainException(
                $"Consumo supera la línea de crédito disponible. Saldo: {Saldo:N2}, Monto: {monto:N2}.",
                "CTACTE_SALDO_INSUFICIENTE");

        Consumo += monto;
    }
}
