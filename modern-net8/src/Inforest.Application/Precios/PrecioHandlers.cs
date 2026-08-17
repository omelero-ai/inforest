using Inforest.Application.Interfaces;
using Inforest.Domain.Common;

namespace Inforest.Application.Precios;

// ──────────────────────────────────────────────────────────────────────────────
// DTOs
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Fila del listado de precios de productos.
/// Legacy: frmPrecios.frm — vProducto (Codigo, Grupo, SubGrupo, Descripcion, nPrecioVenta, nPrecioDelivery, nPrecioLlevar, lActivo)
/// </summary>
public sealed record PrecioProductoFila(
    string Codigo,
    string Grupo,
    string SubGrupo,
    string Descripcion,
    decimal PrecioVenta,
    decimal PrecioDelivery,
    decimal PrecioLlevar,
    bool Activo);

/// <summary>
/// Datos del ítem de pedido para editar precio.
/// Legacy: frmPreciosDetalle.frm — vPedidoDetalle campos relevantes
/// </summary>
public sealed record ItemPedidoParaPrecio(
    string CodigoPedido,
    string Item,
    string CodigoProducto,
    string Producto,
    bool Modificable,
    decimal PrecioOficial,
    decimal PrecioNeto,
    decimal PrecioVenta,
    decimal Descuento,
    decimal Recargo,
    decimal Impuesto1,
    decimal Impuesto2,
    decimal Impuesto3,
    decimal Cantidad,
    bool AplicaExoneracion,
    bool AplicaInafectacion);

// ──────────────────────────────────────────────────────────────────────────────
// Commands / Queries
// ──────────────────────────────────────────────────────────────────────────────

public sealed record ObtenerListadoPreciosQuery;

public sealed record ObtenerItemParaPrecioQuery(string CodigoPedido, string Item);

/// <summary>
/// Modifica el precio de un ítem de pedido.
/// Legacy: frmPreciosDetalle.frm — GrabaProducto: UPDATE DPEDIDO + INSERT TLOG_MODPRECIO
/// BR-PRECIO-001: Solo lModificable=true
/// BR-PRECIO-002: Registra TLOG_MODPRECIO
/// BR-PRECIO-003: Descuento/recargo validado
/// BR-PRECIO-005: Cálculo impuestos Bolivia vs Perú/Ecuador
/// </summary>
public sealed record ModificarPrecioItemCommand(
    string CodigoPedido,
    string Item,
    decimal PrecioNeto,
    decimal PrecioOficial,
    decimal PrecioVenta,
    decimal Descuento,
    decimal Recargo,
    decimal Impuesto1,
    decimal Impuesto2,
    decimal Impuesto3,
    decimal Cantidad,
    string UsuarioAutoriza,
    string Usuario);

// ──────────────────────────────────────────────────────────────────────────────
// Repository interface
// ──────────────────────────────────────────────────────────────────────────────

public interface IPrecioRepository
{
    Task<IEnumerable<PrecioProductoFila>> ObtenerListadoPreciosAsync(CancellationToken ct = default);

    Task<ItemPedidoParaPrecio?> ObtenerItemParaPrecioAsync(string codigoPedido, string item, CancellationToken ct = default);

    Task ModificarPrecioItemAsync(ModificarPrecioItemCommand command, CancellationToken ct = default);
}

// ──────────────────────────────────────────────────────────────────────────────
// Handlers
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Devuelve el listado de precios desde vProducto.
/// Legacy: frmPrecios.frm — Form_Load, SELECT Codigo, Grupo, SubGrupo, Descripcion,
///         nPrecioVenta, nPrecioDelivery, nPrecioLlevar, lActivo FROM vProducto
/// </summary>
public sealed class ObtenerListadoPreciosHandler
{
    private readonly IPrecioRepository _repo;

    public ObtenerListadoPreciosHandler(IPrecioRepository repo) => _repo = repo;

    public async Task<Result<IEnumerable<PrecioProductoFila>>> HandleAsync(
        ObtenerListadoPreciosQuery _, CancellationToken ct = default)
    {
        var filas = await _repo.ObtenerListadoPreciosAsync(ct);
        return Result.Ok(filas);
    }
}

/// <summary>
/// Obtiene los datos de un ítem de pedido para la pantalla de precios.
/// Legacy: frmPreciosDetalle.frm — Form_Load, SELECT * FROM vPedidoDetalle WHERE tEstadoItem='N'
///         AND tCodigoPedido=... AND tItem=...
/// </summary>
public sealed class ObtenerItemParaPrecioHandler
{
    private readonly IPrecioRepository _repo;

    public ObtenerItemParaPrecioHandler(IPrecioRepository repo) => _repo = repo;

    public async Task<Result<ItemPedidoParaPrecio>> HandleAsync(
        ObtenerItemParaPrecioQuery query, CancellationToken ct = default)
    {
        var item = await _repo.ObtenerItemParaPrecioAsync(query.CodigoPedido, query.Item, ct);
        return item is null
            ? Result.Fail<ItemPedidoParaPrecio>($"No se encontró el ítem {query.Item} del pedido {query.CodigoPedido}.", "PRECIO_ITEM_NO_ENCONTRADO")
            : Result.Ok(item);
    }
}

/// <summary>
/// Modifica el precio de un ítem de pedido.
/// Legacy: frmPreciosDetalle.frm — GrabaProducto: UPDATE DPEDIDO + INSERT TLOG_MODPRECIO
/// BR-PRECIO-001: lModificable requerido (validado en Desktop antes de llamar)
/// BR-PRECIO-002: Audit en TLOG_MODPRECIO
/// </summary>
public sealed class ModificarPrecioItemHandler
{
    private readonly IPrecioRepository _repo;

    public ModificarPrecioItemHandler(IPrecioRepository repo) => _repo = repo;

    public async Task<Result> HandleAsync(ModificarPrecioItemCommand command, CancellationToken ct = default)
    {
        try
        {
            await _repo.ModificarPrecioItemAsync(command, ct);
            return Result.Ok();
        }
        catch (Exception ex)
        {
            return Result.Fail($"Error al modificar precio del ítem {command.Item}: {ex.Message}", "PRECIO_ERROR_MODIFICAR");
        }
    }
}
