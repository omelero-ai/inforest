using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Despacho;

// ──────────────────────────────────────────────────────────────────────────────
// DTO
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Fila del Despachador — proyección de <c>vDespachador</c>.
/// <para>
/// Legacy: recordset de <c>vDespachador where tTipoPedido='02' and tEstadoPedido='02'</c>
/// en <c>frmDespachador.frm → cmdProcesa_Click</c>.
/// </para>
/// </summary>
public sealed record FilaDespachador(
    string CodigoPedido,
    string? TipoPedido,
    string? EstadoPedido,
    string? ClienteDelivery,
    string? Telefono,
    string? Direccion,
    string? Zona,
    string? Distrito,
    string? Motorizado,
    DateTime FechaRegistro,
    DateTime? FechaDespacho,
    DateTime? FechaLlegada,
    bool EsEmpacador,
    decimal MontoTotal);

// ──────────────────────────────────────────────────────────────────────────────
// QUERIES
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Query para obtener pedidos del panel despachador.
/// <para>
/// Legacy: <c>vDespachador</c> + filtros en <c>frmDespachador.frm → cmdProcesa_Click</c>.
/// Condición: <c>tTipoPedido='02' AND tEstadoPedido='02' AND fLlegada IS NULL</c>.
/// </para>
/// </summary>
public sealed record ObtenerPedidosDespachadorQuery(
    DateTime FechaInicio,
    DateTime FechaFin,
    string? CodigoCaja = null);

/// <summary>Handler de <see cref="ObtenerPedidosDespachadorQuery"/>.</summary>
public sealed class ObtenerPedidosDespachadorHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ObtenerPedidosDespachadorHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoDelivery>>> HandleAsync(
        ObtenerPedidosDespachadorQuery query, CancellationToken ct = default)
    {
        var pedidos = await _repo.ObtenerParaDespachadorAsync(query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(pedidos);
    }
}

/// <summary>
/// Query para exportar datos del despachador (HTML).
/// <para>
/// Legacy: <c>frmDespachador.frm → cmdExporta_Click</c> — exporta grilla a archivo HTML.
/// </para>
/// </summary>
public sealed record ExportarDespachadorQuery(
    DateTime FechaInicio,
    DateTime FechaFin,
    string? CodigoCaja = null);

/// <summary>Handler de <see cref="ExportarDespachadorQuery"/>.</summary>
public sealed class ExportarDespachadorHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ExportarDespachadorHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoDelivery>>> HandleAsync(
        ExportarDespachadorQuery query, CancellationToken ct = default)
    {
        var pedidos = await _repo.ObtenerParaDespachadorAsync(query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(pedidos);
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// COMMANDS
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para emitir/despachar un pedido.
/// <para>
/// Legacy: <c>frmDespachador.frm → cmdEmite_Click</c> — marca pedido como despachado
/// y genera impresión.
/// </para>
/// </summary>
public sealed record EmitirDespachoPedidoCommand(
    string CodigoPedido,
    string Usuario);

/// <summary>Handler de <see cref="EmitirDespachoPedidoCommand"/>.</summary>
public sealed class EmitirDespachoPedidoHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public EmitirDespachoPedidoHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(EmitirDespachoPedidoCommand cmd, CancellationToken ct = default)
    {
        var pedido = await _repo.ObtenerPorCodigoAsync(cmd.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("No se encontró el pedido.", "DESPACHO_PEDIDO_NO_ENCONTRADO");

        pedido.ActualizarEstado(EstadoDelivery.EnCamino);
        await _repo.ActualizarEstadoAsync(cmd.CodigoPedido, EstadoDelivery.EnCamino, ct);
        return Result.Ok();
    }
}
