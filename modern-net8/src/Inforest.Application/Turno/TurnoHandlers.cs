using Inforest.Application.Configuracion;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Caja;
using TurnoEntity = Inforest.Domain.Entities.Configuracion.Turno;

namespace Inforest.Application.Turno;

public sealed record AbrirTurnoCommand(string CodigoCaja, string CodigoUsuario, decimal MontoInicial);

public sealed class AbrirTurnoHandler
{
    private const bool DiaContableAutomatico = true;
    private const string HoraCierreDiaContable = "05:00";

    private readonly ITurnoRepository _turnoRepository;
    private readonly IDiaContableService _diaContableService;

    public AbrirTurnoHandler(ITurnoRepository turnoRepository, IDiaContableService diaContableService)
    {
        _turnoRepository = turnoRepository;
        _diaContableService = diaContableService;
    }

    public async Task<Result<TurnoEntity>> HandleAsync(AbrirTurnoCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.CodigoCaja))
            return Result.Fail<TurnoEntity>("La caja es obligatoria.", "TURNO_CAJA_REQUERIDA");

        if (string.IsNullOrWhiteSpace(command.CodigoUsuario))
            return Result.Fail<TurnoEntity>("El usuario es obligatorio.", "TURNO_USUARIO_REQUERIDO");

        var turnoActual = await _turnoRepository.ObtenerTurnoActualAsync(command.CodigoCaja, ct);
        if (turnoActual is not null)
            return Result.Fail<TurnoEntity>("La caja ya tiene un turno abierto.", "TURNO_YA_ABIERTO");

        var fechaDiaContable = await _diaContableService.ObtenerDiaContableAsync(
            DiaContableAutomatico,
            HoraCierreDiaContable,
            command.CodigoUsuario,
            ct);

        var codigoTurno = GenerarCodigoTurno();
        var turno = TurnoEntity.Abrir(
            codigoTurno,
            command.CodigoCaja,
            command.CodigoUsuario,
            fechaDiaContable,
            command.MontoInicial);

        var inserted = await _turnoRepository.InsertarAsync(turno, ct);
        return inserted
            ? Result.Ok(turno)
            : Result.Fail<TurnoEntity>("No se pudo aperturar el turno.", "TURNO_APERTURA_FALLIDA");
    }

    private static string GenerarCodigoTurno()
        => DateTime.Now.ToString("yyMMddHHmmss");
}

public sealed record CerrarTurnoCommand(
    string CodigoTurno,
    string CodigoCaja,
    decimal MontoFinal,
    CierreTurnoBreakdown? Breakdown = null,
    bool SupervisorAutorizado = false,
    bool DescargoPendienteConfirmado = false);

/// <summary>
/// Cierra el turno de caja aplicando las reglas de negocio:
/// - BR-CAJA-001: si lObligaCierre=true y no hay autorización de supervisor → retorna REQUIERE_SUPERVISOR
/// - BR-CAJA-002: si lActivaConsultaDescargo=true y no confirmado → retorna REQUIERE_CONFIRMACION_DESCARGO
/// - BR-CAJA-004: actualiza MTURNO con desglose completo de montos.
/// Legacy: frmLiquidacionDetalle.frm, cmdOpcion_Click(0).
/// </summary>
public sealed class CerrarTurnoHandler
{
    private readonly ITurnoRepository _turnoRepository;
    private readonly IParametroRepository _configuracionRepository;

    public CerrarTurnoHandler(
        ITurnoRepository turnoRepository,
        IParametroRepository configuracionRepository)
    {
        _turnoRepository = turnoRepository;
        _configuracionRepository = configuracionRepository;
    }

    public async Task<Result> HandleAsync(CerrarTurnoCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.CodigoTurno))
            return Result.Fail("El código del turno es obligatorio.", "TURNO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(command.CodigoCaja))
            return Result.Fail("El código de caja es obligatorio.", "TURNO_CAJA_REQUERIDA");

        // BR-CAJA-001: lObligaCierre — requiere supervisor autorizado
        var configCaja = await _configuracionRepository.ObtenerConfiguracionCajaAsync(command.CodigoCaja, ct);
        if (configCaja?.lObligaCierre == true && !command.SupervisorAutorizado)
            return Result.Fail("Se requiere autorización de supervisor para cerrar el turno.", "REQUIERE_SUPERVISOR");

        // BR-CAJA-002: lActivaConsultaDescargo — requiere confirmación de que el descargo fue realizado
        var configSistema = await _configuracionRepository.ObtenerConfiguracionAsync(ct);
        if (configSistema?.lActivaConsultaDescargo == true && !command.DescargoPendienteConfirmado)
            return Result.Fail("Debe confirmar que realizó el descargo de ventas antes de cerrar el turno.", "REQUIERE_CONFIRMACION_DESCARGO");

        var breakdown = command.Breakdown ?? CierreTurnoBreakdown.Vacio();
        var closed = await _turnoRepository.CerrarAsync(command.CodigoTurno, command.MontoFinal, breakdown, ct);
        return closed
            ? Result.Ok()
            : Result.Fail("No se pudo cerrar el turno solicitado.", "TURNO_CIERRE_FALLIDO");
    }
}

public sealed record ObtenerTurnoActualQuery(string CodigoCaja);

public sealed class ObtenerTurnoActualHandler
{
    private readonly ITurnoRepository _turnoRepository;

    public ObtenerTurnoActualHandler(ITurnoRepository turnoRepository)
        => _turnoRepository = turnoRepository;

    public async Task<Result<TurnoEntity?>> HandleAsync(ObtenerTurnoActualQuery query, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoCaja))
            return Result.Fail<TurnoEntity?>("La caja es obligatoria.", "TURNO_CAJA_REQUERIDA");

        var turno = await _turnoRepository.ObtenerTurnoActualAsync(query.CodigoCaja, ct);
        return Result.Ok<TurnoEntity?>(turno);
    }
}
