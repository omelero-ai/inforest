using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Delivery;

// ──────────────────────────────────────────────────────────────────────────────
// COMMANDS
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para crear un nuevo cliente de delivery.
/// <para>
/// Legacy: <c>frmNuevoDelivery.frm</c> — acción de guardar cliente frecuente.
/// Tabla: <c>TDELIVERY</c>.
/// </para>
/// </summary>
public sealed record CrearClienteDeliveryCommand(
    string CodigoDelivery,
    string? TipoCliente,
    string? Apellido,
    string? Nombre,
    string? Telefono,
    string? Direccion,
    string? Referencia,
    string? CodigoZona,
    string? CodigoDistrito,
    string? TipoIdentidad,
    string? NumeroIdentidad,
    string? CodigoTarjeta,
    string? NumeroTarjeta,
    string? Observacion,
    string? Email);

/// <summary>
/// Handler de <see cref="CrearClienteDeliveryCommand"/>.
/// </summary>
public sealed class CrearClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public CrearClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(CrearClienteDeliveryCommand cmd, CancellationToken ct = default)
    {
        var existente = await _repo.ObtenerPorCodigoAsync(cmd.CodigoDelivery, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un cliente delivery con ese código.", "DELIVERY_CLIENTE_YA_EXISTE");

        var cliente = ClienteDelivery.Crear(
            cmd.CodigoDelivery,
            cmd.TipoCliente,
            cmd.Apellido,
            cmd.Nombre,
            cmd.Telefono,
            cmd.Direccion,
            cmd.CodigoZona,
            cmd.CodigoDistrito);

        if (!string.IsNullOrWhiteSpace(cmd.TipoIdentidad) && !string.IsNullOrWhiteSpace(cmd.NumeroIdentidad))
            cliente.AsignarIdentidad(cmd.TipoIdentidad, cmd.NumeroIdentidad);

        if (!string.IsNullOrWhiteSpace(cmd.CodigoTarjeta) && !string.IsNullOrWhiteSpace(cmd.NumeroTarjeta))
            cliente.AsignarTarjeta(cmd.CodigoTarjeta, cmd.NumeroTarjeta);

        await _repo.InsertarAsync(cliente, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para actualizar datos de un cliente delivery existente.
/// <para>
/// Legacy: <c>frmNuevoDelivery.frm</c> — edición de cliente frecuente.
/// </para>
/// </summary>
public sealed record ActualizarClienteDeliveryCommand(
    string CodigoDelivery,
    string? Apellido,
    string? Nombre,
    string? Telefono,
    string? Direccion,
    string? Referencia,
    string? CodigoZona,
    string? CodigoDistrito,
    string? Observacion,
    string? Email);

/// <summary>Handler de <see cref="ActualizarClienteDeliveryCommand"/>.</summary>
public sealed class ActualizarClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public ActualizarClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(ActualizarClienteDeliveryCommand cmd, CancellationToken ct = default)
    {
        var cliente = await _repo.ObtenerPorCodigoAsync(cmd.CodigoDelivery, ct);
        if (cliente is null)
            return Result.Fail("No se encontró el cliente delivery.", "DELIVERY_CLIENTE_NO_ENCONTRADO");

        cliente.Actualizar(
            cmd.Apellido, cmd.Nombre, cmd.Telefono, cmd.Direccion,
            cmd.Referencia, cmd.CodigoZona, cmd.CodigoDistrito,
            cmd.Observacion, cmd.Email);

        await _repo.ActualizarAsync(cliente, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para actualizar la foto de un cliente delivery.
/// <para>
/// Legacy: <c>frmClienteDeliveryDetalle.frm → GuardarFoto()</c>.
/// SP: <c>sp_UpdFotoDelivery</c> sobre tabla <c>TDELIVERY.iFoto</c>.
/// </para>
/// </summary>
public sealed record ActualizarFotoClienteDeliveryCommand(
    string CodigoDelivery,
    byte[] Foto);

/// <summary>Handler de <see cref="ActualizarFotoClienteDeliveryCommand"/>.</summary>
public sealed class ActualizarFotoClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public ActualizarFotoClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(ActualizarFotoClienteDeliveryCommand cmd, CancellationToken ct = default)
    {
        var cliente = await _repo.ObtenerPorCodigoAsync(cmd.CodigoDelivery, ct);
        if (cliente is null)
            return Result.Fail("No se encontró el cliente delivery.", "DELIVERY_CLIENTE_NO_ENCONTRADO");

        if (cmd.Foto.Length == 0)
            return Result.Fail("La foto no puede estar vacía.", "DELIVERY_FOTO_VACIA");

        await _repo.ActualizarFotoAsync(cmd.CodigoDelivery, cmd.Foto, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para crear un pedido delivery.
/// <para>
/// Legacy: <c>frmNuevoDelivery.frm</c> + <c>frmPedido.frm</c> en módulo Despachador.
/// Tabla: <c>MPEDIDO</c> con campos de delivery.
/// </para>
/// Reglas: BR-DEL-001, BR-DEL-010.
/// </summary>
public sealed record CrearPedidoDeliveryCommand(
    string CodigoPedido,
    string CodigoTurno,
    string CodigoUsuario,
    string CodigoCaja,
    string CodigoClienteDelivery,
    string? CodigoZona,
    string? CodigoDistrito,
    string? TipoPedido,
    DateTime? HoraEstimadaEntrega);

/// <summary>Handler de <see cref="CrearPedidoDeliveryCommand"/>.</summary>
public sealed class CrearPedidoDeliveryHandler
{
    private readonly IPedidoDeliveryRepository _pedidoRepo;

    public CrearPedidoDeliveryHandler(IPedidoDeliveryRepository pedidoRepo)
        => _pedidoRepo = pedidoRepo;

    public async Task<Result> HandleAsync(CrearPedidoDeliveryCommand cmd, CancellationToken ct = default)
    {
        var pedido = PedidoDelivery.Crear(
            cmd.CodigoPedido,
            cmd.CodigoTurno,
            cmd.CodigoUsuario,
            cmd.CodigoCaja,
            cmd.CodigoClienteDelivery,
            cmd.CodigoZona,
            cmd.CodigoDistrito,
            cmd.TipoPedido);

        if (cmd.HoraEstimadaEntrega.HasValue)
            pedido.EstablecerHoraEstimada(cmd.HoraEstimadaEntrega.Value);

        await _pedidoRepo.InsertarAsync(pedido, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para actualizar el estado de un pedido delivery.
/// <para>
/// Legacy: encapsula <c>sp_CD_Modificar_EstadoDelivery_Cabecera</c> en CENTRALDELIVERY.
/// </para>
/// Regla BR-DEL-009.
/// </summary>
public sealed record ActualizarEstadoDeliveryCommand(
    string CodigoPedido,
    EstadoDelivery NuevoEstado,
    string Usuario);

/// <summary>Handler de <see cref="ActualizarEstadoDeliveryCommand"/>.</summary>
public sealed class ActualizarEstadoDeliveryHandler
{
    private readonly IPedidoDeliveryRepository _pedidoRepo;
    private readonly ICentralPedidosRepository _centralRepo;

    public ActualizarEstadoDeliveryHandler(
        IPedidoDeliveryRepository pedidoRepo,
        ICentralPedidosRepository centralRepo)
    {
        _pedidoRepo = pedidoRepo;
        _centralRepo = centralRepo;
    }

    public async Task<Result> HandleAsync(ActualizarEstadoDeliveryCommand cmd, CancellationToken ct = default)
    {
        var pedido = await _pedidoRepo.ObtenerPorCodigoAsync(cmd.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("No se encontró el pedido delivery.", "DELIVERY_PEDIDO_NO_ENCONTRADO");

        pedido.ActualizarEstado(cmd.NuevoEstado);
        await _pedidoRepo.ActualizarEstadoAsync(cmd.CodigoPedido, cmd.NuevoEstado, ct);

        // Propagar al Central de Pedidos si la conexión está disponible (BR-DEL-003, BR-DEL-009)
        if (await _centralRepo.VerificarConexionAsync(ct))
        {
            await _centralRepo.ModificarEstadoDeliveryAsync(
                cmd.CodigoPedido,
                cmd.NuevoEstado.ToString(),
                cmd.Usuario,
                ct);
        }

        return Result.Ok();
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// QUERIES
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Query para buscar clientes de delivery.
/// <para>
/// Legacy: vista <c>vDelivery</c> con filtros en <c>frmClienteDelivery.frm</c>.
/// </para>
/// </summary>
public sealed record BuscarClienteDeliveryQuery(
    string? Apellido = null,
    string? Nombre = null,
    string? Telefono = null,
    string? CodigoDistrito = null);

/// <summary>Handler de <see cref="BuscarClienteDeliveryQuery"/>.</summary>
public sealed class BuscarClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public BuscarClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<ClienteDelivery>>> HandleAsync(
        BuscarClienteDeliveryQuery query, CancellationToken ct = default)
    {
        var clientes = await _repo.BuscarAsync(query.Apellido, query.Nombre, query.Telefono, query.CodigoDistrito, ct);
        return Result.Ok(clientes);
    }
}

/// <summary>
/// Query para obtener clientes frecuentes.
/// <para>
/// Legacy: <c>frmClienteDeliveryFrecuente.frm</c>.
/// </para>
/// </summary>
public sealed record ObtenerClientesFrecuentesQuery(int Top = 20);

/// <summary>Handler de <see cref="ObtenerClientesFrecuentesQuery"/>.</summary>
public sealed class ObtenerClientesFrecuentesHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public ObtenerClientesFrecuentesHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<ClienteDelivery>>> HandleAsync(
        ObtenerClientesFrecuentesQuery query, CancellationToken ct = default)
    {
        var clientes = await _repo.ObtenerFrecuentesAsync(query.Top, ct);
        return Result.Ok(clientes);
    }
}

/// <summary>
/// Query para obtener pedidos delivery pendientes.
/// <para>
/// Legacy: vista <c>vDelivery</c> con filtros de fecha/estado/zona en <c>frmNuevoDelivery.frm</c>.
/// </para>
/// </summary>
public sealed record ObtenerDeliveryPendientesQuery(
    string CodigoCaja,
    DateTime FechaInicio,
    DateTime FechaFin);

/// <summary>Handler de <see cref="ObtenerDeliveryPendientesQuery"/>.</summary>
public sealed class ObtenerDeliveryPendientesHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ObtenerDeliveryPendientesHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoDelivery>>> HandleAsync(
        ObtenerDeliveryPendientesQuery query, CancellationToken ct = default)
    {
        var pedidos = await _repo.ObtenerPendientesAsync(query.CodigoCaja, query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(pedidos);
    }
}
