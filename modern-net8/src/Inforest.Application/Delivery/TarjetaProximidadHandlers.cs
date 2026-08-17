using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Delivery;

/// <summary>
/// Contrato de acceso a datos para tarjetas RFID/proximidad.
/// Legacy: <c>TTARJETASRFID</c> + <c>TMOVIMIENTOTARJETASRFID</c>.
/// </summary>
public interface ITarjetaProximidadRepository
{
    Task<IReadOnlyList<TarjetaProximidad>> ObtenerTodasAsync(CancellationToken ct = default);
    Task<TarjetaProximidad?> ObtenerPorCodigoAsync(string codigoRfid, CancellationToken ct = default);
    Task<bool> InsertarAsync(TarjetaProximidad tarjeta, CancellationToken ct = default);
    Task<bool> ActualizarAsync(TarjetaProximidad tarjeta, CancellationToken ct = default);
    Task<IReadOnlyList<MovimientoTarjetaProximidad>> ObtenerUltimosMovimientosAsync(string codigoRfid, int cantidad = 10, CancellationToken ct = default);
}

public sealed record ObtenerTarjetasProximidadQuery;

public sealed record ObtenerMovimientosTarjetaProximidadQuery(string CodigoRfid, int Cantidad = 10);

public sealed record CrearTarjetaProximidadCommand(
    string CodigoRfid,
    string Descripcion,
    decimal MontoDisponible,
    string CodigoCliente,
    string Estado,
    string Usuario);

public sealed record ActualizarTarjetaProximidadCommand(
    string CodigoRfid,
    string Descripcion,
    decimal MontoDisponible,
    string CodigoCliente,
    string Estado,
    string Usuario);

/// <summary>
/// Lista tarjetas RFID con el cliente asociado.
/// Legacy: <c>FrmTarjetaAproximidad.Form_Load</c>.
/// </summary>
public sealed class ObtenerTarjetasProximidadHandler
{
    private readonly ITarjetaProximidadRepository _repository;

    public ObtenerTarjetasProximidadHandler(ITarjetaProximidadRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<TarjetaProximidad>>> HandleAsync(ObtenerTarjetasProximidadQuery query, CancellationToken ct = default)
    {
        var tarjetas = await _repository.ObtenerTodasAsync(ct);
        return Result.Ok<IReadOnlyList<TarjetaProximidad>>(tarjetas);
    }
}

/// <summary>
/// Obtiene los últimos movimientos de una tarjeta para consulta/impresión.
/// Legacy: <c>FrmTarjetaAproximidad.Genera</c>.
/// </summary>
public sealed class ObtenerMovimientosTarjetaProximidadHandler
{
    private readonly ITarjetaProximidadRepository _repository;

    public ObtenerMovimientosTarjetaProximidadHandler(ITarjetaProximidadRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<MovimientoTarjetaProximidad>>> HandleAsync(ObtenerMovimientosTarjetaProximidadQuery query, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoRfid))
            return Result.Fail<IReadOnlyList<MovimientoTarjetaProximidad>>(
                "Debe seleccionar una tarjeta.",
                "RFID_CODIGO_REQUERIDO");

        var movimientos = await _repository.ObtenerUltimosMovimientosAsync(query.CodigoRfid.Trim(), Math.Max(query.Cantidad, 1), ct);
        return Result.Ok<IReadOnlyList<MovimientoTarjetaProximidad>>(movimientos);
    }
}

/// <summary>
/// Crea una tarjeta RFID nueva validando código único y cliente existente.
/// Legacy: <c>FrmTarjetaAproximidadDetalle.cmdOpcion_Click(Index=1)</c>.
/// </summary>
public sealed class CrearTarjetaProximidadHandler
{
    private readonly ITarjetaProximidadRepository _repository;
    private readonly IClienteDeliveryRepository _clienteDeliveryRepository;

    public CrearTarjetaProximidadHandler(
        ITarjetaProximidadRepository repository,
        IClienteDeliveryRepository clienteDeliveryRepository)
    {
        _repository = repository;
        _clienteDeliveryRepository = clienteDeliveryRepository;
    }

    public async Task<Result> HandleAsync(CrearTarjetaProximidadCommand command, CancellationToken ct = default)
    {
        try
        {
            var existente = await _repository.ObtenerPorCodigoAsync(command.CodigoRfid.Trim(), ct);
            if (existente is not null)
                return Result.Fail("Código de tarjeta ya existe.", "RFID_CODIGO_DUPLICADO");

            var cliente = await _clienteDeliveryRepository.ObtenerPorCodigoAsync(command.CodigoCliente.Trim(), ct);
            if (cliente is null)
                return Result.Fail("El cliente asociado no existe.", "RFID_CLIENTE_NO_EXISTE");

            var tarjeta = TarjetaProximidad.Crear(
                command.CodigoRfid,
                command.Descripcion,
                command.MontoDisponible,
                command.CodigoCliente,
                command.Estado,
                command.Usuario);

            tarjeta.AsignarNombreCliente($"{cliente.Apellido} {cliente.Nombre}".Trim());

            var creado = await _repository.InsertarAsync(tarjeta, ct);
            return creado
                ? Result.Ok()
                : Result.Fail("No se pudo registrar la tarjeta.", "RFID_INSERT_FALLO");
        }
        catch (DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
    }
}

/// <summary>
/// Actualiza una tarjeta RFID existente validando cliente y estado permitido.
/// Legacy: <c>FrmTarjetaAproximidadDetalle.cmdOpcion_Click(Index=1)</c>.
/// </summary>
public sealed class ActualizarTarjetaProximidadHandler
{
    private readonly ITarjetaProximidadRepository _repository;
    private readonly IClienteDeliveryRepository _clienteDeliveryRepository;

    public ActualizarTarjetaProximidadHandler(
        ITarjetaProximidadRepository repository,
        IClienteDeliveryRepository clienteDeliveryRepository)
    {
        _repository = repository;
        _clienteDeliveryRepository = clienteDeliveryRepository;
    }

    public async Task<Result> HandleAsync(ActualizarTarjetaProximidadCommand command, CancellationToken ct = default)
    {
        try
        {
            var tarjeta = await _repository.ObtenerPorCodigoAsync(command.CodigoRfid.Trim(), ct);
            if (tarjeta is null)
                return Result.Fail("No se encontró la tarjeta a modificar.", "RFID_NO_ENCONTRADA");

            var cliente = await _clienteDeliveryRepository.ObtenerPorCodigoAsync(command.CodigoCliente.Trim(), ct);
            if (cliente is null)
                return Result.Fail("El cliente asociado no existe.", "RFID_CLIENTE_NO_EXISTE");

            tarjeta.Actualizar(
                command.Descripcion,
                command.MontoDisponible,
                command.CodigoCliente,
                command.Estado,
                command.Usuario);

            tarjeta.AsignarNombreCliente($"{cliente.Apellido} {cliente.Nombre}".Trim());

            var actualizado = await _repository.ActualizarAsync(tarjeta, ct);
            return actualizado
                ? Result.Ok()
                : Result.Fail("No se pudo actualizar la tarjeta.", "RFID_UPDATE_FALLO");
        }
        catch (DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
    }
}
