using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Configuracion;

/// <summary>
/// Turno de caja — equivalente a tabla MTURNO.
/// Legacy: MTURNO en 1. Estructura.sql, clsDiaContable.cls
/// Regla BR-004: El turno determina el período de operación de una caja.
/// </summary>
public class Turno : Entity
{
    public string CodigoTurno { get; private set; } = string.Empty;
    public string CodigoCaja { get; private set; } = string.Empty;
    public string CodigoSalon { get; private set; } = string.Empty;
    public string CodigoUsuario { get; private set; } = string.Empty;
    public DateTime FechaApertura { get; private set; }
    public DateTime? FechaCierre { get; private set; }
    public DateTime FechaDiaContable { get; private set; }
    /// <summary>nMontoIN — Monto inicial en moneda nacional.</summary>
    public decimal MontoInicial { get; private set; }
    /// <summary>nMontoIE — Monto inicial en moneda extranjera.</summary>
    public decimal MontoInicialME { get; private set; }
    public decimal MontoFinal { get; private set; }
    public bool Cerrado { get; private set; }

    private Turno() { }

    /// <summary>
    /// Crea una nueva apertura de turno.
    /// Legacy: frmInicio.frm cmdOpcion_Click(0) — INSERT INTO MTURNO.
    /// BR-TURNO-001.
    /// </summary>
    public static Turno Abrir(
        string codigoTurno,
        string codigoCaja,
        string codigoUsuario,
        DateTime fechaDiaContable,
        decimal montoInicial = 0,
        decimal montoInicialME = 0,
        string? codigoSalon = null)
    {
        if (string.IsNullOrWhiteSpace(codigoTurno))
            throw new DomainException("El código de turno es obligatorio.", "TURNO_CODIGO_REQUERIDO");

        return new Turno
        {
            CodigoTurno = codigoTurno,
            CodigoCaja = codigoCaja,
            CodigoSalon = codigoSalon ?? string.Empty,
            CodigoUsuario = codigoUsuario,
            FechaApertura = DateTime.Now,
            FechaDiaContable = fechaDiaContable,
            MontoInicial = montoInicial,
            MontoInicialME = montoInicialME,
            Cerrado = false
        };
    }

    public void Cerrar(decimal montoFinal)
    {
        if (Cerrado)
            throw new DomainException("El turno ya está cerrado.", "TURNO_YA_CERRADO");

        MontoFinal = montoFinal;
        FechaCierre = DateTime.Now;
        Cerrado = true;
    }
}
