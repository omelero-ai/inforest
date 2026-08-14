using Inforest.Domain.Common;
using Inforest.Domain.Entities.Cocina;
using Inforest.Domain.Exceptions;

namespace Inforest.Application.Kitchen;

public sealed record ObtenerMensajesCocinaQuery(DateTime FechaInicio, DateTime FechaFin, string CodigoCaja = "");
public sealed record AgregarMensajeCocinaCommand(string Mensaje, string Usuario, string CodigoCaja, bool Activo, bool ActivarAlerta);
public sealed record ModificarMensajeCocinaCommand(string Codigo, string Mensaje, string Usuario, string CodigoCaja, bool Activo, bool ActivarAlerta);
public sealed record EliminarMensajeCocinaCommand(string Codigo);
public sealed record CerrarMensajesCocinaPorCajaCommand(string Usuario, string CodigoCaja);

/// <summary>
/// Lista mensajes de cocina por rango/ caja heredando el comportamiento del SP legacy.
/// Legacy: frmMensajeCocina.frm + USP_LISTARMENSAJES.
/// Regla BR-MSGCOC-004.
/// </summary>
public sealed class ObtenerMensajesCocinaHandler
{
    private readonly IMensajeCocinaRepository _repository;

    public ObtenerMensajesCocinaHandler(IMensajeCocinaRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<MensajeCocina>>> HandleAsync(ObtenerMensajesCocinaQuery query, CancellationToken ct = default)
    {
        if (query.FechaInicio.Date > query.FechaFin.Date)
            return Result.Fail<IReadOnlyList<MensajeCocina>>("Error en rango de fechas.", "MENSAJE_COCINA_RANGO_INVALIDO");

        var mensajes = await _repository.ObtenerMensajesAsync(query.FechaInicio, query.FechaFin, query.CodigoCaja ?? string.Empty, ct);
        return Result.Ok(mensajes);
    }
}

/// <summary>
/// Registra un mensaje de cocina respetando correlativo anual y límite de mensajes activos.
/// Legacy: frmMensajeCocinaDetalle.frm + USP_AGREGARMENSAJE.
/// Reglas BR-MSGCOC-001, BR-MSGCOC-002, BR-MSGCOC-003.
/// </summary>
public sealed class AgregarMensajeCocinaHandler
{
    private readonly IMensajeCocinaRepository _repository;

    public AgregarMensajeCocinaHandler(IMensajeCocinaRepository repository)
        => _repository = repository;

    public async Task<Result<string>> HandleAsync(AgregarMensajeCocinaCommand command, CancellationToken ct = default)
    {
        try
        {
            if (command.Activo && await _repository.ContarActivosAsync(ct: ct) >= MensajeCocina.MaximoMensajesActivos)
                return Result.Fail<string>(
                    $"Solo puede tener máximo {MensajeCocina.MaximoMensajesActivos} mensajes a cocina activos.",
                    "MENSAJE_COCINA_MAXIMO_ACTIVOS");

            var codigo = await _repository.ObtenerProximoCodigoAsync(DateTime.Today, ct);
            var mensaje = MensajeCocina.Crear(
                codigo,
                command.Usuario,
                command.Mensaje,
                command.CodigoCaja,
                command.Activo,
                command.ActivarAlerta);

            await _repository.AgregarAsync(mensaje, ct);
            return Result.Ok(codigo);
        }
        catch (DomainException ex)
        {
            return Result.Fail<string>(ex.Message, ex.Code);
        }
    }
}

/// <summary>
/// Modifica un mensaje de cocina existente.
/// Legacy: frmMensajeCocinaDetalle.frm + USP_MODIFICARMENSAJE.
/// Reglas BR-MSGCOC-001, BR-MSGCOC-003.
/// </summary>
public sealed class ModificarMensajeCocinaHandler
{
    private readonly IMensajeCocinaRepository _repository;

    public ModificarMensajeCocinaHandler(IMensajeCocinaRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(ModificarMensajeCocinaCommand command, CancellationToken ct = default)
    {
        try
        {
            var mensaje = await _repository.ObtenerPorCodigoAsync(command.Codigo, ct);
            if (mensaje is null)
                return Result.Fail("No se encontró el mensaje a modificar.", "MENSAJE_COCINA_NO_ENCONTRADO");

            if (command.Activo && await _repository.ContarActivosAsync(command.Codigo, ct) >= MensajeCocina.MaximoMensajesActivos)
                return Result.Fail(
                    $"Solo puede tener máximo {MensajeCocina.MaximoMensajesActivos} mensajes a cocina activos.",
                    "MENSAJE_COCINA_MAXIMO_ACTIVOS");

            mensaje.Actualizar(command.Usuario, command.Mensaje, command.CodigoCaja, command.Activo, command.ActivarAlerta);
            await _repository.ModificarAsync(mensaje, ct);
            return Result.Ok();
        }
        catch (DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
    }
}

/// <summary>
/// Elimina un mensaje de cocina.
/// Legacy: frmMensajeCocinaDetalle.frm + USP_ELIMINARRMENSAJES.
/// Regla BR-MSGCOC-004.
/// </summary>
public sealed class EliminarMensajeCocinaHandler
{
    private readonly IMensajeCocinaRepository _repository;

    public EliminarMensajeCocinaHandler(IMensajeCocinaRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(EliminarMensajeCocinaCommand command, CancellationToken ct = default)
    {
        var mensaje = await _repository.ObtenerPorCodigoAsync(command.Codigo, ct);
        if (mensaje is null)
            return Result.Fail("No se encontró el mensaje a eliminar.", "MENSAJE_COCINA_NO_ENCONTRADO");

        await _repository.EliminarAsync(command.Codigo, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Cierra mensajes activos de la caja al finalizar turno.
/// Legacy: frmLiquidacionDetalle.frm + USP_CERRAR_MENSAJES_CIERRETURNO.
/// Regla BR-MSGCOC-005.
/// </summary>
public sealed class CerrarMensajesCocinaPorCajaHandler
{
    private readonly IMensajeCocinaRepository _repository;

    public CerrarMensajesCocinaPorCajaHandler(IMensajeCocinaRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(CerrarMensajesCocinaPorCajaCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.Usuario))
            return Result.Fail("El usuario es obligatorio.", "MENSAJE_COCINA_USUARIO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(command.CodigoCaja))
            return Result.Fail("La caja es obligatoria.", "MENSAJE_COCINA_CAJA_REQUERIDA");

        await _repository.CerrarActivosPorCajaAsync(command.Usuario, command.CodigoCaja, ct);
        return Result.Ok();
    }
}
