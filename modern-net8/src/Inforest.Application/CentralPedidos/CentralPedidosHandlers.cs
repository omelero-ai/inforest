using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Application.CentralPedidos;

// ──────────────────────────────────────────────────────────────────────────────
// QUERIES
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Query para obtener pedidos de la Central de Pedidos multi-local.
/// <para>
/// Legacy SP: <c>sp_Inforest_PedidosCentralPedido</c> en base <c>CENTRALDELIVERY</c>.
/// Formulario: <c>frmCentralPedidos.frm</c>.
/// </para>
/// Reglas: BR-DEL-003, BR-DEL-008.
/// </summary>
public sealed record ObtenerPedidosCentralQuery(
    DateTime FechaInicio,
    DateTime FechaFin);

/// <summary>Handler de <see cref="ObtenerPedidosCentralQuery"/>.</summary>
public sealed class ObtenerPedidosCentralHandler
{
    private readonly ICentralPedidosRepository _repo;

    public ObtenerPedidosCentralHandler(ICentralPedidosRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoCentral>>> HandleAsync(
        ObtenerPedidosCentralQuery query, CancellationToken ct = default)
    {
        if (!await _repo.VerificarConexionAsync(ct))
            return Result.Fail<IEnumerable<PedidoCentral>>(
                "Central de Pedidos no disponible. Verificar conexión a CENTRALDELIVERY.",
                "CENTRAL_PEDIDOS_NO_DISPONIBLE");

        var pedidos = await _repo.ObtenerPedidosCentralAsync(query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(pedidos);
    }
}

/// <summary>
/// Query para obtener locales activos (multi-local).
/// <para>
/// Legacy: tabla <c>TLOCAL</c> filtrada por <c>lActivo = 1</c>.
/// </para>
/// Regla BR-DEL-007.
/// </summary>
public sealed record ObtenerLocalesActivosQuery;

/// <summary>Handler de <see cref="ObtenerLocalesActivosQuery"/>.</summary>
public sealed class ObtenerLocalesActivosHandler
{
    private readonly ILocalRepository _repo;

    public ObtenerLocalesActivosHandler(ILocalRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<LocalRestaurante>>> HandleAsync(
        ObtenerLocalesActivosQuery _, CancellationToken ct = default)
    {
        var locales = await _repo.ObtenerActivosAsync(ct);
        return Result.Ok(locales);
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// COMMANDS
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para sincronizar/replicar un pedido hacia CENTRALDELIVERY.
/// <para>
/// Legacy: conexión directa a <c>CENTRALDELIVERY</c> desde <c>modDespachador.bas</c>.
/// </para>
/// Regla BR-DEL-008.
/// </summary>
public sealed record SincronizarPedidoCentralCommand(
    string CodigoPedido,
    string CodigoLocal,
    string Usuario);

/// <summary>Handler de <see cref="SincronizarPedidoCentralCommand"/>.</summary>
public sealed class SincronizarPedidoCentralHandler
{
    private readonly ICentralPedidosRepository _centralRepo;
    private readonly IPedidoDeliveryRepository _pedidoRepo;

    public SincronizarPedidoCentralHandler(
        ICentralPedidosRepository centralRepo,
        IPedidoDeliveryRepository pedidoRepo)
    {
        _centralRepo = centralRepo;
        _pedidoRepo = pedidoRepo;
    }

    public async Task<Result> HandleAsync(SincronizarPedidoCentralCommand cmd, CancellationToken ct = default)
    {
        if (!await _centralRepo.VerificarConexionAsync(ct))
            return Result.Fail("Central de Pedidos no disponible.", "CENTRAL_PEDIDOS_NO_DISPONIBLE");

        var pedido = await _pedidoRepo.ObtenerPorCodigoAsync(cmd.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("Pedido no encontrado.", "PEDIDO_NO_ENCONTRADO");

        // Sincroniza estado actual hacia CENTRALDELIVERY
        await _centralRepo.ModificarEstadoDeliveryAsync(
            cmd.CodigoPedido,
            pedido.EstadoDelivery.ToString(),
            cmd.Usuario,
            ct);

        return Result.Ok();
    }
}
