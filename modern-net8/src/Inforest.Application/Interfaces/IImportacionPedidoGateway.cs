using Inforest.Application.Almacen;
using Inforest.Domain.Common;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Gateway para crear un pedido en INFOREST a partir de un requerimiento de almacén.
/// Legacy: InsertaProducto() + spIns_MPEDIDO en frmImportacionRequerimientos.frm.
/// Aísla la lógica de precio/impuesto por canal y la inserción en DPEDIDO.
/// BR-IMPORT-003, BR-IMPORT-004.
/// </summary>
public interface IImportacionPedidoGateway
{
    /// <summary>
    /// Crea el pedido completo (MPEDIDO + DPEDIDO) a partir de un contexto de importación.
    /// Si algún producto no existe en INFOREST, cancela el pedido (estado '03').
    /// </summary>
    Task<Result<ImportarRequerimientoResult>> CrearPedidoDesdeRequerimientoAsync(
        ImportacionPedidoContexto contexto,
        CancellationToken ct = default);
}

/// <summary>
/// Datos de contexto necesarios para crear un pedido desde un requerimiento.
/// </summary>
public sealed record ImportacionPedidoContexto(
    string Rq,
    string CodigoCaja,
    string CodigoTurno,
    string CodigoSalon,
    string CodigoUsuario,
    string TipoPedido,
    DateTime FechaDiaContable,
    string Observacion,
    IReadOnlyList<ItemImportacionPedido> Detalle);

/// <summary>
/// Ítem de pedido a crear desde un requerimiento.
/// Legacy: enlace vía TPRODUCTO.tEnlace → vProducto.Codigo.
/// </summary>
public sealed record ItemImportacionPedido(
    string CodigoProductoInforRest,
    double Cantidad);
