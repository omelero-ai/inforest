using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Exceptions;

namespace Inforest.Application.Delivery;

/// <summary>
/// Movimiento de recarga RFID.
/// Legacy: FrmRecargarTarjeta.frm + usp_Inforest_ObtieneRecargas + TMOVIMIENTOTARJETASRFID.
/// Reglas: BR-RFID-005.
/// </summary>
public sealed record RecargaTarjetaMovimiento(
    int Id,
    string CodigoRfid,
    string Tipo,
    string DocumentoReferencia,
    DateTime FechaRegistro,
    decimal MontoIngreso,
    decimal MontoSalida,
    decimal MontoAnterior,
    decimal MontoFinal);

/// <summary>
/// Resultado de registro de recarga de tarjeta.
/// </summary>
public sealed record RecargaTarjetaResultado(
    string CodigoRfid,
    string DocumentoReferencia,
    decimal MontoAnterior,
    decimal MontoRecarga,
    decimal MontoFinal);

public interface IRecargaTarjetaRepository
{
    Task<IReadOnlyList<RecargaTarjetaMovimiento>> ObtenerMovimientosAsync(
        DateTime desde,
        DateTime hasta,
        string tipoMovimiento,
        string? codigoRfid = null,
        CancellationToken ct = default);

    Task<RecargaTarjetaResultado> RegistrarRecargaAsync(
        string codigoRfid,
        decimal montoRecarga,
        string documentoReferencia,
        CancellationToken ct = default);
}

public sealed record ObtenerRecargasTarjetaQuery(
    DateTime Desde,
    DateTime Hasta,
    string TipoMovimiento = "R",
    string? CodigoRfid = null);

public sealed class ObtenerRecargasTarjetaHandler
{
    private readonly IRecargaTarjetaRepository _repository;

    public ObtenerRecargasTarjetaHandler(IRecargaTarjetaRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<RecargaTarjetaMovimiento>>> HandleAsync(
        ObtenerRecargasTarjetaQuery query,
        CancellationToken ct = default)
    {
        if (query.Desde > query.Hasta)
            return Result.Fail<IReadOnlyList<RecargaTarjetaMovimiento>>(
                "La fecha de inicio no puede ser mayor a la fecha fin.",
                "RFID_RECARGA_RANGO_INVALIDO");

        var tipo = string.IsNullOrWhiteSpace(query.TipoMovimiento)
            ? "R"
            : query.TipoMovimiento.Trim().ToUpperInvariant();

        var items = await _repository.ObtenerMovimientosAsync(
            query.Desde,
            query.Hasta,
            tipo,
            string.IsNullOrWhiteSpace(query.CodigoRfid) ? null : query.CodigoRfid.Trim(),
            ct);

        return Result.Ok<IReadOnlyList<RecargaTarjetaMovimiento>>(items);
    }
}

public sealed record RegistrarRecargaTarjetaCommand(
    string CodigoRfid,
    decimal MontoRecarga,
    string Usuario,
    string? DocumentoReferencia = null);

public sealed class RegistrarRecargaTarjetaHandler
{
    private readonly ITarjetaProximidadRepository _tarjetaRepository;
    private readonly IRecargaTarjetaRepository _recargaRepository;

    public RegistrarRecargaTarjetaHandler(
        ITarjetaProximidadRepository tarjetaRepository,
        IRecargaTarjetaRepository recargaRepository)
    {
        _tarjetaRepository = tarjetaRepository;
        _recargaRepository = recargaRepository;
    }

    public async Task<Result<RecargaTarjetaResultado>> HandleAsync(
        RegistrarRecargaTarjetaCommand command,
        CancellationToken ct = default)
    {
        var codigoRfid = (command.CodigoRfid ?? string.Empty).Trim();
        if (codigoRfid.Length == 0)
            return Result.Fail<RecargaTarjetaResultado>(
                "Debe seleccionar una tarjeta.",
                "RFID_RECARGA_TARJETA_REQUERIDA");

        if (command.MontoRecarga <= 0)
            return Result.Fail<RecargaTarjetaResultado>(
                "El monto de recarga debe ser mayor a cero.",
                "RFID_RECARGA_MONTO_INVALIDO");

        var usuario = (command.Usuario ?? string.Empty).Trim();
        if (usuario.Length == 0)
            return Result.Fail<RecargaTarjetaResultado>(
                "El usuario es obligatorio para registrar la recarga.",
                "RFID_RECARGA_USUARIO_REQUERIDO");

        var tarjeta = await _tarjetaRepository.ObtenerPorCodigoAsync(codigoRfid, ct);
        if (tarjeta is null)
            return Result.Fail<RecargaTarjetaResultado>(
                "La tarjeta indicada no existe.",
                "RFID_RECARGA_TARJETA_NO_EXISTE");

        if (string.Equals(tarjeta.Estado, "Bloqueado", StringComparison.OrdinalIgnoreCase))
            return Result.Fail<RecargaTarjetaResultado>(
                "No se puede recargar una tarjeta bloqueada.",
                "RFID_RECARGA_TARJETA_BLOQUEADA");

        try
        {
            var documentoReferencia = string.IsNullOrWhiteSpace(command.DocumentoReferencia)
                ? $"REC-{DateTime.Now:yyyyMMddHHmmss}"
                : command.DocumentoReferencia.Trim();

            var resultado = await _recargaRepository.RegistrarRecargaAsync(
                codigoRfid,
                decimal.Round(command.MontoRecarga, 2, MidpointRounding.AwayFromZero),
                documentoReferencia,
                ct);

            return Result.Ok(resultado);
        }
        catch (DomainException ex)
        {
            return Result.Fail<RecargaTarjetaResultado>(ex.Message, ex.Code);
        }
    }
}
