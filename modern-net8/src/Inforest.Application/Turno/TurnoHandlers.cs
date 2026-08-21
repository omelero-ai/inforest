using Inforest.Application.Configuracion;
using Inforest.Application.Kitchen;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Caja;
using TurnoEntity = Inforest.Domain.Entities.Configuracion.Turno;

namespace Inforest.Application.Turno;

/// <summary>
/// Comando de apertura de turno.
/// Legacy: frmInicio.frm cmdOpcion_Click(0) — INSERT/UPDATE MTURNO.
/// BR-TURNO-001 (apertura nueva), BR-TURNO-002 (re-apertura).
/// </summary>
public sealed record AbrirTurnoCommand(
    string CodigoCaja,
    string CodigoUsuario,
    decimal MontoInicial,
    decimal MontoInicialME = 0m,
    string? CodigoSalon = null,
    ModoConsultaTurno ModoTurno = ModoConsultaTurno.PorCaja,
    bool RegistrarTipoCambio = false,
    decimal TipoCambio = 0m,
    decimal TipoCambioOficial = 0m,
    decimal TipoCambio2 = 0m,
    decimal TipoCambio3 = 0m
);

/// <summary>
/// Resultado de apertura de turno, incluye el código generado y si fue re-apertura.
/// </summary>
public sealed record AbrirTurnoResult(TurnoEntity Turno, bool EsReApertura, string CodigoTurno);

/// <summary>
/// Apertura de turno de caja.
///
/// Legacy: frmInicio.frm cmdOpcion_Click(0).
///
/// Flujo:
///   1. Si existe turno con lCierre=0 → re-apertura: UPDATE MTURNO SET tUsuario, nMontoIN, nMontoIE (BR-TURNO-002).
///   2. Si no existe turno o el último tiene lCierre=1 → apertura nueva:
///      a. Genera correlativo: YY + Correlativo(MAX(tTurno) últimos 8, 8).
///      b. INSERT INTO MTURNO (tTurno, tCaja, tSalon, fInicial, tUsuario, lCierre, nMontoIN, nMontoIE).
///      (BR-TURNO-001)
///   3. Si se indicó RegistrarTipoCambio → spIns_TipoCambio.
///
/// Reglas: BR-TURNO-001, BR-TURNO-002, BR-TC-001.
/// </summary>
public sealed class AbrirTurnoHandler
{
    private readonly ITurnoRepository _turnoRepository;
    private readonly IDiaContableService _diaContableService;
    private readonly ITipoCambioRepository? _tipoCambioRepository;

    public AbrirTurnoHandler(
        ITurnoRepository turnoRepository,
        IDiaContableService diaContableService,
        ITipoCambioRepository? tipoCambioRepository = null)
    {
        _turnoRepository = turnoRepository;
        _diaContableService = diaContableService;
        _tipoCambioRepository = tipoCambioRepository;
    }

    public async Task<Result<AbrirTurnoResult>> HandleAsync(AbrirTurnoCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.CodigoCaja))
            return Result.Fail<AbrirTurnoResult>("La caja es obligatoria.", "TURNO_CAJA_REQUERIDA");

        if (string.IsNullOrWhiteSpace(command.CodigoUsuario))
            return Result.Fail<AbrirTurnoResult>("El usuario es obligatorio.", "TURNO_USUARIO_REQUERIDO");

        // Registrar tipo de cambio si fue indicado (spIns_TipoCambio)
        // BR-TC-001: se inserta/actualiza antes de aperturar el turno.
        if (command.RegistrarTipoCambio && _tipoCambioRepository is not null)
        {
            var tc = new TipoCambioDelDia(command.TipoCambio, command.TipoCambioOficial, command.TipoCambio2, command.TipoCambio3);
            await _tipoCambioRepository.InsertarOActualizarAsync(tc, command.CodigoUsuario, ct);
        }

        // Verificar si existe turno para la caja (cualquier estado)
        var ultimo = await _turnoRepository.ObtenerUltimoTurnoAsync(
            command.CodigoCaja,
            command.CodigoUsuario,
            command.ModoTurno,
            ct);

        bool esReApertura = ultimo is not null && !ultimo.Cerrado;

        if (esReApertura)
        {
            // BR-TURNO-002: re-apertura — UPDATE MTURNO SET tUsuario, nMontoIN, nMontoIE
            var updated = await _turnoRepository.ReAperturarAsync(
                ultimo!.CodigoTurno,
                command.CodigoUsuario,
                command.MontoInicial,
                command.MontoInicialME,
                ct);

            if (!updated)
                return Result.Fail<AbrirTurnoResult>("No se pudo re-aperturar el turno.", "TURNO_REAPERTURA_FALLIDA");

            // Reconstruir entidad para el resultado
            var fechaDC = await _diaContableService.ObtenerDiaContableAsync(true, "05:00", command.CodigoUsuario, ct);
            var turnoRe = TurnoEntity.Abrir(ultimo.CodigoTurno, command.CodigoCaja, command.CodigoUsuario, fechaDC, command.MontoInicial);
            return Result.Ok(new AbrirTurnoResult(turnoRe, true, ultimo.CodigoTurno));
        }

        // BR-TURNO-001: apertura nueva
        var fechaDiaContable = await _diaContableService.ObtenerDiaContableAsync(true, "05:00", command.CodigoUsuario, ct);
        var codigoTurno = await _turnoRepository.GenerarCorrelativoAsync(ct);
        var turnoNuevo = TurnoEntity.Abrir(
            codigoTurno,
            command.CodigoCaja,
            command.CodigoUsuario,
            fechaDiaContable,
            command.MontoInicial,
            command.MontoInicialME,
            command.CodigoSalon);

        var inserted = await _turnoRepository.InsertarAsync(turnoNuevo, ct);
        return inserted
            ? Result.Ok(new AbrirTurnoResult(turnoNuevo, false, codigoTurno))
            : Result.Fail<AbrirTurnoResult>("No se pudo aperturar el turno.", "TURNO_APERTURA_FALLIDA");
    }
}

public sealed record CerrarTurnoCommand(
    string CodigoTurno,
    string CodigoCaja,
    decimal MontoFinal,
    CierreTurnoBreakdown? Breakdown = null,
    bool SupervisorAutorizado = false,
    bool DescargoPendienteConfirmado = false,
    string? CodigoUsuario = null);

/// <summary>
/// Cierra el turno de caja aplicando las reglas de negocio:
/// - BR-CAJA-001: si lObligaCierre=true y no hay autorización de supervisor → retorna REQUIERE_SUPERVISOR
/// - BR-CAJA-002: si lActivaConsultaDescargo=true y no confirmado → retorna REQUIERE_CONFIRMACION_DESCARGO
/// - BR-CAJA-004: actualiza MTURNO con desglose completo de montos.
/// - BR-MSGCOC-005: cierra mensajes activos de cocina para la caja al finalizar el turno.
/// Legacy: frmLiquidacionDetalle.frm, cmdOpcion_Click(0).
/// </summary>
public sealed class CerrarTurnoHandler
{
    private readonly ITurnoRepository _turnoRepository;
    private readonly IParametroRepository _configuracionRepository;
    private readonly IMensajeCocinaRepository? _mensajeCocinaRepository;

    public CerrarTurnoHandler(
        ITurnoRepository turnoRepository,
        IParametroRepository configuracionRepository,
        IMensajeCocinaRepository? mensajeCocinaRepository = null)
    {
        _turnoRepository = turnoRepository;
        _configuracionRepository = configuracionRepository;
        _mensajeCocinaRepository = mensajeCocinaRepository;
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
        if (!closed)
            return Result.Fail("No se pudo cerrar el turno solicitado.", "TURNO_CIERRE_FALLIDO");

        if (_mensajeCocinaRepository is not null && !string.IsNullOrWhiteSpace(command.CodigoUsuario))
            await _mensajeCocinaRepository.CerrarActivosPorCajaAsync(command.CodigoUsuario, command.CodigoCaja, ct);

        return Result.Ok();
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

public sealed record ObtenerUltimoTurnoQuery(
    string CodigoCaja,
    string CodigoUsuario,
    ModoConsultaTurno Modo = ModoConsultaTurno.PorCaja);

/// <summary>
/// Obtiene el último turno de la caja (cualquier estado, incluyendo cerrados).
/// Legacy: frmInicio.frm Form_Load — SELECT * FROM MTURNO WHERE ... ORDER BY tTurno; RsTurno.MoveLast.
/// Usado por FrmAperturaTurno para determinar si es apertura nueva o re-apertura y cargar montos anteriores.
/// </summary>
public sealed class ObtenerUltimoTurnoHandler
{
    private readonly ITurnoRepository _turnoRepository;

    public ObtenerUltimoTurnoHandler(ITurnoRepository turnoRepository)
        => _turnoRepository = turnoRepository;

    public async Task<Result<TurnoExistente?>> HandleAsync(ObtenerUltimoTurnoQuery query, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoCaja))
            return Result.Fail<TurnoExistente?>("La caja es obligatoria.", "TURNO_CAJA_REQUERIDA");

        var ultimo = await _turnoRepository.ObtenerUltimoTurnoAsync(query.CodigoCaja, query.CodigoUsuario, query.Modo, ct);
        return Result.Ok<TurnoExistente?>(ultimo);
    }
}

public sealed record ValidarInicioCajaRapidaQuery(
    string CodigoCaja,
    string CodigoUsuario,
    ModoConsultaTurno Modo = ModoConsultaTurno.PorCaja);

public sealed record ValidarInicioCajaRapidaResult(
    bool PermiteIngresoDirecto,
    string CodigoTurno = "");

/// <summary>
/// Valida si el POS multi-cajero puede entrar directo sin mostrar la apertura de turno.
/// Legacy: <c>mdiPuntoVenta.frm</c> → <c>validaInicioCajaRapida()</c>.
/// Requiere tipo de cambio del día y un turno abierto según el modo consultado.
/// Si ambos existen, reaplica usuario/montos al turno abierto y habilita el inicio.
/// </summary>
public sealed class ValidarInicioCajaRapidaHandler
{
    private readonly ITurnoRepository _turnoRepository;
    private readonly ITipoCambioRepository _tipoCambioRepository;

    public ValidarInicioCajaRapidaHandler(
        ITurnoRepository turnoRepository,
        ITipoCambioRepository tipoCambioRepository)
    {
        _turnoRepository = turnoRepository;
        _tipoCambioRepository = tipoCambioRepository;
    }

    public async Task<Result<ValidarInicioCajaRapidaResult>> HandleAsync(ValidarInicioCajaRapidaQuery query, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoCaja))
            return Result.Fail<ValidarInicioCajaRapidaResult>("La caja es obligatoria.", "TURNO_CAJA_REQUERIDA");

        if (string.IsNullOrWhiteSpace(query.CodigoUsuario))
            return Result.Fail<ValidarInicioCajaRapidaResult>("El usuario es obligatorio.", "TURNO_USUARIO_REQUERIDO");

        var tipoCambio = await _tipoCambioRepository.ObtenerDelDiaAsync(ct);
        if (tipoCambio is null || tipoCambio.Venta <= 0)
            return Result.Ok(new ValidarInicioCajaRapidaResult(false));

        var ultimo = await _turnoRepository.ObtenerUltimoTurnoAsync(
            query.CodigoCaja,
            query.CodigoUsuario,
            query.Modo,
            ct);

        if (ultimo is null || ultimo.Cerrado)
            return Result.Ok(new ValidarInicioCajaRapidaResult(false));

        var actualizado = await _turnoRepository.ReAperturarAsync(
            ultimo.CodigoTurno,
            query.CodigoUsuario,
            ultimo.MontoInicialMN,
            ultimo.MontoInicialME,
            ct);

        if (!actualizado)
            return Result.Fail<ValidarInicioCajaRapidaResult>(
                "No se pudo reutilizar el turno activo para el inicio directo.",
                "TURNO_INICIO_DIRECTO_FALLIDO");

        return Result.Ok(new ValidarInicioCajaRapidaResult(true, ultimo.CodigoTurno));
    }
}
