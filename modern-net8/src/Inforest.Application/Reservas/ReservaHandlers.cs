using Inforest.Application.Interfaces;
using Inforest.Application.Turno;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Reservas;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Reservas;

// ──────────────────────────────────────────────────────────────────────────────
// Commands / Queries
// ──────────────────────────────────────────────────────────────────────────────

public sealed record CrearReservaCommand(
    DateTime FechaReserva,
    string HoraReserva,
    string? CodigoCliente,
    string? NombreCliente,
    string? Telefono,
    string? CodigoMotivo,
    string? CodigoMesero,
    int Pax,
    string? CodigoMesa,
    string? Observacion);

public sealed record ModificarReservaCommand(
    string CodigoReserva,
    DateTime FechaReserva,
    string HoraReserva,
    string? CodigoCliente,
    string? NombreCliente,
    string? Telefono,
    string? CodigoMotivo,
    string? CodigoMesero,
    int Pax,
    string? CodigoMesa,
    string? Observacion);

public sealed record AnularReservaCommand(string CodigoReserva);

public sealed record ObtenerReservaQuery(string CodigoReserva);

public sealed record ObtenerReservasPorFechaQuery(DateTime FechaInicio, DateTime FechaFin);

public sealed record ConvertirReservaAPedidoCommand(
    string CodigoReserva,
    string CodigoCaja,
    string CodigoUsuario,
    string CodigoTurno,
    DateTime FechaDiaContable);

// ──────────────────────────────────────────────────────────────────────────────
// Handlers
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Crea una nueva reserva.
/// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 1 (Sw=True, INSERT INTO TRESERVA)
/// BR-RESERVA-001, BR-RESERVA-002, BR-RESERVA-003
/// </summary>
public sealed class CrearReservaHandler
{
    private readonly IReservaRepository _reservaRepository;
    private readonly IDiaContableService _diaContableService;

    public CrearReservaHandler(IReservaRepository reservaRepository, IDiaContableService diaContableService)
    {
        _reservaRepository = reservaRepository;
        _diaContableService = diaContableService;
    }

    public async Task<Result<Reserva>> HandleAsync(CrearReservaCommand command, CancellationToken ct = default)
    {
        try
        {
            var codigo = await _reservaRepository.GenerarProximoCodigoAsync(ct);
            var diaContable = await _diaContableService.ObtenerFechaServidorAsync(ct);

            var reserva = Reserva.Crear(
                codigo,
                command.FechaReserva,
                command.HoraReserva,
                command.CodigoCliente,
                command.NombreCliente,
                command.Telefono,
                command.CodigoMotivo,
                command.CodigoMesero,
                command.Pax,
                command.CodigoMesa,
                command.Observacion,
                diaContable);

            await _reservaRepository.InsertarAsync(reserva, ct);
            return Result.Ok(reserva);
        }
        catch (Domain.Exceptions.DomainException ex)
        {
            return Result.Fail<Reserva>(ex.Message, ex.Code);
        }
    }
}

/// <summary>
/// Modifica una reserva existente.
/// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 1 (Sw=False, UPDATE TRESERVA)
/// BR-RESERVA-004
/// </summary>
public sealed class ModificarReservaHandler
{
    private readonly IReservaRepository _reservaRepository;

    public ModificarReservaHandler(IReservaRepository reservaRepository)
        => _reservaRepository = reservaRepository;

    public async Task<Result> HandleAsync(ModificarReservaCommand command, CancellationToken ct = default)
    {
        try
        {
            var reserva = await _reservaRepository.ObtenerPorCodigoAsync(command.CodigoReserva, ct);
            if (reserva is null)
                return Result.Fail($"No se encontró la reserva {command.CodigoReserva}.", "RESERVA_NO_ENCONTRADA");

            reserva.Modificar(
                command.FechaReserva,
                command.HoraReserva,
                command.CodigoCliente,
                command.NombreCliente,
                command.Telefono,
                command.CodigoMotivo,
                command.CodigoMesero,
                command.Pax,
                command.CodigoMesa,
                command.Observacion);

            await _reservaRepository.ActualizarAsync(reserva, ct);
            return Result.Ok();
        }
        catch (Domain.Exceptions.DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
    }
}

/// <summary>
/// Anula una reserva.
/// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 2 (UPDATE tEstadoReserva='03')
/// BR-RESERVA-004
/// </summary>
public sealed class AnularReservaHandler
{
    private readonly IReservaRepository _reservaRepository;

    public AnularReservaHandler(IReservaRepository reservaRepository)
        => _reservaRepository = reservaRepository;

    public async Task<Result> HandleAsync(AnularReservaCommand command, CancellationToken ct = default)
    {
        try
        {
            var reserva = await _reservaRepository.ObtenerPorCodigoAsync(command.CodigoReserva, ct);
            if (reserva is null)
                return Result.Fail($"No se encontró la reserva {command.CodigoReserva}.", "RESERVA_NO_ENCONTRADA");

            reserva.Anular();
            await _reservaRepository.ActualizarAsync(reserva, ct);
            return Result.Ok();
        }
        catch (Domain.Exceptions.DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
    }
}

/// <summary>
/// Obtiene una reserva por código.
/// </summary>
public sealed class ObtenerReservaHandler
{
    private readonly IReservaRepository _reservaRepository;

    public ObtenerReservaHandler(IReservaRepository reservaRepository)
        => _reservaRepository = reservaRepository;

    public async Task<Result<Reserva>> HandleAsync(ObtenerReservaQuery query, CancellationToken ct = default)
    {
        var reserva = await _reservaRepository.ObtenerPorCodigoAsync(query.CodigoReserva, ct);
        return reserva is null
            ? Result.Fail<Reserva>($"No se encontró la reserva {query.CodigoReserva}.", "RESERVA_NO_ENCONTRADA")
            : Result.Ok(reserva);
    }
}

/// <summary>
/// Obtiene reservas en un rango de fechas.
/// Legacy: frmReserva.frm — Form_Load / cmdProcesa_Click (SELECT FROM TRESERVA WHERE fFecha between ...)
/// </summary>
public sealed class ObtenerReservasPorFechaHandler
{
    private readonly IReservaRepository _reservaRepository;

    public ObtenerReservasPorFechaHandler(IReservaRepository reservaRepository)
        => _reservaRepository = reservaRepository;

    public async Task<Result<IEnumerable<Reserva>>> HandleAsync(ObtenerReservasPorFechaQuery query, CancellationToken ct = default)
    {
        var reservas = await _reservaRepository.ObtenerPorFechaAsync(query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(reservas);
    }
}

/// <summary>
/// Convierte una reserva pendiente en un pedido activo.
/// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 4 (spIns_MPEDIDO_RESERVA)
/// BR-RESERVA-004: Solo Pendiente puede convertirse
/// </summary>
public sealed class ConvertirReservaAPedidoHandler
{
    private readonly IReservaRepository _reservaRepository;

    public ConvertirReservaAPedidoHandler(IReservaRepository reservaRepository)
        => _reservaRepository = reservaRepository;

    public async Task<Result<string>> HandleAsync(ConvertirReservaAPedidoCommand command, CancellationToken ct = default)
    {
        try
        {
            var reserva = await _reservaRepository.ObtenerPorCodigoAsync(command.CodigoReserva, ct);
            if (reserva is null)
                return Result.Fail<string>($"No se encontró la reserva {command.CodigoReserva}.", "RESERVA_NO_ENCONTRADA");

            // BR-RESERVA-004: solo reservas pendientes
            reserva.MarcarAtendida();

            // spIns_MPEDIDO_RESERVA delegado a Infrastructure (requiere DynamicParameters/Dapper output)
            var codigoPedido = await _reservaRepository.ConvertirAPedidoAsync(
                command.CodigoReserva,
                command.CodigoCaja,
                command.CodigoUsuario,
                command.CodigoTurno,
                command.FechaDiaContable,
                ct);

            if (string.IsNullOrWhiteSpace(codigoPedido))
                return Result.Fail<string>($"Error al generar pedido desde la reserva {command.CodigoReserva}.", "RESERVA_SP_SIN_PEDIDO");

            // Persistir el estado Atendida en TRESERVA
            await _reservaRepository.ActualizarAsync(reserva, ct);

            return Result.Ok(codigoPedido);
        }
        catch (Domain.Exceptions.DomainException ex)
        {
            return Result.Fail<string>(ex.Message, ex.Code);
        }
    }
}
