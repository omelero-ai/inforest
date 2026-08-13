using Inforest.Domain.Common;
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

public sealed record CerrarTurnoCommand(string CodigoTurno, decimal MontoFinal);

public sealed class CerrarTurnoHandler
{
    private readonly ITurnoRepository _turnoRepository;

    public CerrarTurnoHandler(ITurnoRepository turnoRepository)
        => _turnoRepository = turnoRepository;

    public async Task<Result> HandleAsync(CerrarTurnoCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.CodigoTurno))
            return Result.Fail("El código del turno es obligatorio.", "TURNO_CODIGO_REQUERIDO");

        var closed = await _turnoRepository.CerrarAsync(command.CodigoTurno, command.MontoFinal, ct);
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
