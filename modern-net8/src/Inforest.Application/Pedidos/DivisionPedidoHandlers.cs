using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Pedidos;

// ─── Commands / Queries ────────────────────────────────────────────────────────

/// <summary>
/// Comando para iniciar una sesión de división de pedido.
/// </summary>
public sealed record IniciarDivisionCommand(string CodigoPedido, string CodigoCaja, string Correlativo, string PrefijioObservacion);

/// <summary>
/// Comando para confirmar y persistir la división.
/// </summary>
public sealed record ConfirmarDivisionCommand(SesionDivision Sesion, string FechaDiaContable);

/// <summary>
/// Comando para agregar un pedido destino a la sesión.
/// </summary>
public sealed record AgregarPedidoDestinoCommand(SesionDivision Sesion, string Correlativo, string Observacion);

/// <summary>
/// Comando para disgregar un item por su código tItem.
/// </summary>
public sealed record DisgregarItemCommand(SesionDivision Sesion, string TItem, double? CantidadPrimera);

/// <summary>
/// Comando para mover un item al pedido destino.
/// </summary>
public sealed record MoverItemCommand(SesionDivision Sesion, string TItem, string CodigoPedidoDestino, bool ToDestino);

// ─── IniciarDivisionHandler ────────────────────────────────────────────────────

/// <summary>
/// Carga los items del pedido origen y crea una SesionDivision en memoria.
/// Legacy: Form_Load de frmDivision.frm — INSERT en tablas temporales + queries.
/// </summary>
public sealed class IniciarDivisionHandler
{
    private readonly IDivisionPedidoRepository _repo;

    public IniciarDivisionHandler(IDivisionPedidoRepository repo)
        => _repo = repo;

    public async Task<Result<SesionDivision>> HandleAsync(
        IniciarDivisionCommand command,
        CancellationToken cancellationToken = default)
    {
        // BR-DIV-001: validar que el pedido esté en estado '01'
        var estado = await _repo.ObtenerEstadoPedidoAsync(command.CodigoPedido, cancellationToken);
        if (estado != "01")
            return Result<SesionDivision>.Fail(
                $"El pedido {command.CodigoPedido} no puede dividirse: estado '{estado ?? "desconocido"}' diferente de Emitido.");

        var tipoPedido = await _repo.ObtenerTipoPedidoAsync(command.CodigoPedido, cancellationToken);
        var records = await _repo.CargarItemsPedidoAsync(command.CodigoPedido, cancellationToken);

        var items = records.Select(r => new ItemDivision(
            item: r.Item,
            codigoPedido: r.CodigoPedido,
            tipoPedido: r.TipoPedido,
            codigoProducto: r.CodigoProducto,
            codigoGrupo: r.CodigoGrupo,
            codigoSubGrupo: r.CodigoSubGrupo,
            moneda: r.Moneda,
            precioNeto: r.PrecioNeto,
            precioImpuesto1: r.PrecioImpuesto1,
            precioImpuesto2: r.PrecioImpuesto2,
            precioImpuesto3: r.PrecioImpuesto3,
            precioVenta: r.PrecioVenta,
            recargo: r.Recargo,
            descuento: r.Descuento,
            precioOficial: r.PrecioOficial,
            cantidad: r.Cantidad,
            impuesto1: r.Impuesto1,
            impuesto2: r.Impuesto2,
            impuesto3: r.Impuesto3,
            venta: r.Venta,
            observacion: r.Observacion,
            cortesia: r.Cortesia,
            imprime: r.Imprime,
            estadoItem: r.EstadoItem,
            area: r.Area)
        {
            NombreProducto = r.NombreProducto
        }).ToList();

        var sesion = SesionDivision.Crear(command.CodigoPedido, tipoPedido, items);
        return Result<SesionDivision>.Ok(sesion);
    }
}

// ─── DisgregarItemHandler ──────────────────────────────────────────────────────

/// <summary>
/// Disgrega un item con cantidad > 1 en unidades individuales (o dos bloques).
/// Legacy: cmdDetalle_Click Index=0 en frmDivision.frm.
/// BR-DIV-003.
/// </summary>
public sealed class DisgregarItemHandler
{
    public Result HandleAsync(DisgregarItemCommand command)
    {
        if (command.CantidadPrimera.HasValue &&
            (command.CantidadPrimera.Value <= 0 || command.CantidadPrimera.Value >= GetCantidad(command.Sesion, command.TItem)))
            return Result.Fail("Cantidad no aceptada.");

        bool ok = command.Sesion.Disgregar(command.TItem, command.CantidadPrimera);
        return ok ? Result.Ok() : Result.Fail($"No se puede disgregar el item '{command.TItem}'.");
    }

    private static double GetCantidad(SesionDivision sesion, string tItem)
    {
        var item = sesion.Items.FirstOrDefault(i => i.Item == tItem);
        return item?.Cantidad ?? 0;
    }
}

// ─── CompartirHandler ──────────────────────────────────────────────────────────

/// <summary>
/// Comparte los items del origen equitativamente entre todos los pedidos.
/// Legacy: cmdDetalle_Click Index=1 en frmDivision.frm.
/// BR-DIV-004.
/// </summary>
public sealed class CompartirPedidoHandler
{
    public Result HandleAsync(SesionDivision sesion)
    {
        if (sesion.PedidosDestino.Count == 0)
            return Result.Fail("Debe agregar al menos un pedido destino antes de compartir.");

        bool ok = sesion.Compartir();
        return ok ? Result.Ok() : Result.Fail("No se puede compartir el pedido.");
    }
}

// ─── MoverItemHandler ─────────────────────────────────────────────────────────

/// <summary>
/// Mueve un item entre el pedido origen y un pedido destino (o viceversa).
/// Legacy: cmdMovimiento_Click en frmDivision.frm.
/// BR-DIV-005.
/// </summary>
public sealed class MoverItemHandler
{
    public Result HandleAsync(MoverItemCommand command)
    {
        bool ok = command.ToDestino
            ? command.Sesion.MoverItemADestino(command.TItem, command.CodigoPedidoDestino)
            : command.Sesion.RevertirItemAOrigen(command.TItem, command.CodigoPedidoDestino);

        return ok ? Result.Ok() : Result.Fail($"No se puede mover el item '{command.TItem}'.");
    }
}

// ─── ConfirmarDivisionHandler ──────────────────────────────────────────────────

/// <summary>
/// Persiste la división de pedido en la base de datos.
/// Legacy: cmdOpcion_Click Index=0 (Aceptar) en frmDivision.frm.
/// BR-DIV-001, BR-DIV-002.
/// </summary>
public sealed class ConfirmarDivisionHandler
{
    private readonly IDivisionPedidoRepository _repo;

    public ConfirmarDivisionHandler(IDivisionPedidoRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(
        ConfirmarDivisionCommand command,
        CancellationToken cancellationToken = default)
    {
        var sesion = command.Sesion;

        // BR-DIV-001: re-validar estado pedido al momento de confirmar
        var estado = await _repo.ObtenerEstadoPedidoAsync(sesion.CodigoPedidoOrigen, cancellationToken);
        if (estado != "01")
            return Result.Fail(
                $"El pedido {sesion.CodigoPedidoOrigen} no puede ser Dividido, Estado del pedido Diferente de Emitido.");

        // BR-DIV-002: validar montoMaximo para cada pedido destino con items
        foreach (var destino in sesion.PedidosDestino)
        {
            var items = sesion.ItemsDestino(destino.CodigoTemporal).ToList();
            if (items.Count == 0) continue;

            var montoMax = await _repo.ObtenerMontoMaximoPedidoAsync(destino.CodigoTemporal, cancellationToken);
            if (montoMax > 0)
            {
                decimal ventaTemp = (decimal)items.Sum(i => i.Venta);
                var ventaActual = await _repo.ObtenerVentaActualPedidoAsync(destino.CodigoTemporal, cancellationToken);
                if (ventaActual + ventaTemp > montoMax)
                    return Result.Fail(
                        $"El pedido {destino.CodigoTemporal} supera el monto máximo de {montoMax:N2}.");
            }
        }

        // Generar códigos definitivos para pedidos temporales (3 chars)
        foreach (var destino in sesion.PedidosDestino)
        {
            var itemsDestino = sesion.ItemsDestino(destino.CodigoTemporal).ToList();
            if (destino.CodigoTemporal.Length == 3 && itemsDestino.Count > 0)
            {
                var codigoDefinitivo = await _repo.GenerarSiguienteCorrelativoPedidoAsync(cancellationToken);
                sesion.AsignarCodigoDefinitivoPedido(destino.CodigoTemporal, codigoDefinitivo);
            }
        }

        var pedidosConItems = sesion.PedidosDestino
            .Select(p => (
                Pedido: p,
                Items: (IReadOnlyList<ItemDivision>)sesion.ItemsDestino(p.CodigoTemporal).ToList()))
            .Where(x => x.Items.Count > 0)
            .ToList();

        var parametros = new ConfirmarDivisionParams(
            CodigoPedidoOrigen: sesion.CodigoPedidoOrigen,
            TipoPedido: sesion.TipoPedido,
            FechaDiaContable: command.FechaDiaContable,
            ItemsOrigen: sesion.ItemsOrigen.ToList(),
            PedidosDestino: pedidosConItems,
            TieneDisgregados: sesion.Items.Any(i => i.TienePropiedad && i.CodigoPedidoDestino is null));

        await _repo.ConfirmarDivisionAsync(parametros, cancellationToken);
        return Result.Ok();
    }
}
