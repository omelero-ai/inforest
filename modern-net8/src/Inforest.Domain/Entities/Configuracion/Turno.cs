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
    public string CodigoUsuario { get; private set; } = string.Empty;
    public DateTime FechaApertura { get; private set; }
    public DateTime? FechaCierre { get; private set; }
    public DateTime FechaDiaContable { get; private set; }
    public decimal MontoInicial { get; private set; }
    public decimal MontoFinal { get; private set; }
    public bool Cerrado { get; private set; }

    private Turno() { }

    public static Turno Abrir(
        string codigoTurno,
        string codigoCaja,
        string codigoUsuario,
        DateTime fechaDiaContable,
        decimal montoInicial = 0)
    {
        if (string.IsNullOrWhiteSpace(codigoTurno))
            throw new DomainException("El código de turno es obligatorio.", "TURNO_CODIGO_REQUERIDO");

        return new Turno
        {
            CodigoTurno = codigoTurno,
            CodigoCaja = codigoCaja,
            CodigoUsuario = codigoUsuario,
            FechaApertura = DateTime.Now,
            FechaDiaContable = fechaDiaContable,
            MontoInicial = montoInicial,
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
