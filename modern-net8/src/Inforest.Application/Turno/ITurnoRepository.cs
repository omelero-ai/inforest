using Inforest.Domain.Entities.Caja;
using TurnoCaja = Inforest.Domain.Entities.Configuracion.Turno;

namespace Inforest.Application.Turno;

public interface ITurnoRepository
{
    /// <summary>
    /// Obtiene el turno activo (lCierre=0) de la caja.
    /// Legacy: SELECT * FROM MTURNO WHERE tCaja = sCaja AND lCierre = 0
    /// </summary>
    Task<TurnoCaja?> ObtenerTurnoActualAsync(string codigoCaja, CancellationToken ct = default);

    /// <summary>
    /// Obtiene el último turno de la caja (cualquier estado).
    /// Legacy: SELECT * FROM MTURNO WHERE tCaja = sCaja ORDER BY tTurno — RsTurno.MoveLast.
    /// Modo MCPV: filtra por usuario. Modo TurnoCompartido: filtra por caja+usuario.
    /// </summary>
    Task<TurnoExistente?> ObtenerUltimoTurnoAsync(
        string codigoCaja,
        string codigoUsuario,
        ModoConsultaTurno modo,
        CancellationToken ct = default);

    /// <summary>
    /// Genera el siguiente código correlativo de turno.
    /// Legacy: SELECT MAX(tTurno) FROM MTURNO WHERE SUBSTRING(tTurno,1,2) = YY
    /// Formato: YY + 8 dígitos secuenciales (ej. "2600000001").
    /// </summary>
    Task<string> GenerarCorrelativoAsync(CancellationToken ct = default);

    /// <summary>
    /// Inserta un nuevo turno.
    /// Legacy: INSERT INTO MTURNO(tTurno, tCaja, tSalon, fInicial, tUsuario, lCierre, nMontoIN, nMontoIE)
    /// BR-TURNO-001.
    /// </summary>
    Task<bool> InsertarAsync(TurnoCaja turno, CancellationToken ct = default);

    /// <summary>
    /// Re-apertura de un turno existente no cerrado: actualiza usuario y montos iniciales.
    /// Legacy: UPDATE MTURNO SET tUsuario, nMontoIN, nMontoIE WHERE tTurno.
    /// BR-TURNO-002.
    /// </summary>
    Task<bool> ReAperturarAsync(string codigoTurno, string codigoUsuario, decimal montoInicial, decimal montoInicialME, CancellationToken ct = default);

    /// <summary>
    /// Cierra el turno actualizando MTURNO con lCierre=1 y el desglose completo de montos.
    /// Legacy: UPDATE MTURNO SET lCierre=1, fFinal=getdate(), nMontoEN=..., nMontoFN=..., etc.
    /// BR-CAJA-004.
    /// </summary>
    Task<bool> CerrarAsync(string codigoTurno, decimal montoFinal, CierreTurnoBreakdown breakdown, CancellationToken ct = default);

    Task<IReadOnlyList<TurnoCaja>> ObtenerHistorialAsync(string codigoCaja, DateTime desde, CancellationToken ct = default);
}

/// <summary>
/// Resultado al consultar el último turno de una caja (cualquier estado).
/// Equivale al RecordSet RsTurno de frmInicio.frm Form_Load.
/// </summary>
public sealed record TurnoExistente(
    string CodigoTurno,
    string CodigoCaja,
    string CodigoUsuario,
    DateTime FechaInicial,
    bool Cerrado,
    decimal MontoInicialMN,
    decimal MontoInicialME
);

/// <summary>
/// Modo de consulta de turno activo.
/// Legacy: lMCPV → por usuario; lTurnoCompartido → caja+usuario; normal → caja.
/// </summary>
public enum ModoConsultaTurno
{
    PorCaja,
    PorUsuario,          // lMCPV
    PorCajaYUsuario      // lTurnoCompartido
}
