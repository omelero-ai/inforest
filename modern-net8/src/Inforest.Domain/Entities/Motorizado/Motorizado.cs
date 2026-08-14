using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Motorizado;

/// <summary>
/// Tipo de día para cálculo de tarifa de motorizado.
/// <para>
/// Legacy: <c>TMOTORIZADODATOS</c> — campos nTarifaLV, nTarifaSD, nTarifaES.
/// </para>
/// Regla BR-DEL-002.
/// </summary>
public enum TipoDiaTarifa
{
    /// <summary>Lunes a Viernes.</summary>
    LunesViernes,
    /// <summary>Sábado y Domingo.</summary>
    SabadoDomingo,
    /// <summary>Día especial (feriado u otros).</summary>
    Especial
}

/// <summary>
/// Datos del motorizado de delivery.
/// <para>
/// Legacy: tabla <c>TMOTORIZADODATOS</c> en <c>INFOREST</c>.
/// Vista asociada: <c>vMotorizado</c>.
/// Formularios: <c>frmMotorizado.frm</c>, <c>frmMotorizadoDetalle.frm</c>,
///   <c>frmAsignacionMotorizado.frm</c>, <c>frmReasignacionMotorizado.frm</c>.
/// </para>
/// Reglas: BR-DEL-002, BR-DEL-005, BR-DEL-006.
/// </summary>
public class Motorizado : Entity
{
    public string Codigo { get; private set; } = string.Empty;
    public string DocumentoIdentidad { get; private set; } = string.Empty;

    /// <summary>Tarifa lunes a viernes. Legacy: <c>nTarifaLV</c>.</summary>
    public decimal TarifaLunesViernes { get; private set; }

    /// <summary>Tarifa sábado y domingo. Legacy: <c>nTarifaSD</c>.</summary>
    public decimal TarifaSabadoDomingo { get; private set; }

    /// <summary>Tarifa días especiales. Legacy: <c>nTarifaES</c>.</summary>
    public decimal TarifaEspecial { get; private set; }

    // Dapper constructor
    private Motorizado() { }

    public static Motorizado Crear(
        string codigo,
        string documentoIdentidad,
        decimal tarifaLunesViernes,
        decimal tarifaSabadoDomingo,
        decimal tarifaEspecial)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código de motorizado es obligatorio.", "MOTORIZADO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(documentoIdentidad))
            throw new DomainException("El documento de identidad del motorizado es obligatorio.", "MOTORIZADO_DOCUMENTO_REQUERIDO");

        if (tarifaLunesViernes < 0 || tarifaSabadoDomingo < 0 || tarifaEspecial < 0)
            throw new DomainException("Las tarifas del motorizado no pueden ser negativas.", "MOTORIZADO_TARIFA_NEGATIVA");

        return new Motorizado
        {
            Codigo = codigo,
            DocumentoIdentidad = documentoIdentidad,
            TarifaLunesViernes = tarifaLunesViernes,
            TarifaSabadoDomingo = tarifaSabadoDomingo,
            TarifaEspecial = tarifaEspecial
        };
    }

    /// <summary>
    /// Obtiene la tarifa aplicable según el tipo de día.
    /// Regla BR-DEL-002.
    /// </summary>
    public decimal ObtenerTarifaPorDia(TipoDiaTarifa tipoDia) => tipoDia switch
    {
        TipoDiaTarifa.LunesViernes => TarifaLunesViernes,
        TipoDiaTarifa.SabadoDomingo => TarifaSabadoDomingo,
        TipoDiaTarifa.Especial => TarifaEspecial,
        _ => throw new DomainException($"Tipo de día de tarifa no reconocido: {tipoDia}.", "MOTORIZADO_TIPO_DIA_INVALIDO")
    };

    public void ActualizarTarifas(decimal lv, decimal sd, decimal especial)
    {
        if (lv < 0 || sd < 0 || especial < 0)
            throw new DomainException("Las tarifas del motorizado no pueden ser negativas.", "MOTORIZADO_TARIFA_NEGATIVA");

        TarifaLunesViernes = lv;
        TarifaSabadoDomingo = sd;
        TarifaEspecial = especial;
    }
}
