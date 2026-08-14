using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Cocina;

/// <summary>
/// Ítem de producción derivado de DPEDIDO/CPEDIDO para comandas y KDS.
/// Legacy: frmVenta.frm + modKDS.bas + TPRODUCTOAREA + DPEDIDOKDS.
/// Reglas BR-010 y BR-011.
/// </summary>
public sealed class ItemProduccion
{
    public string CodigoPedido { get; }
    public string Item { get; }
    public string? ItemCombo { get; }
    public string CodigoProducto { get; }
    public string NombreProducto { get; }
    public decimal Cantidad { get; }
    public string? CodigoArea { get; }
    public string? EstacionKds { get; }
    public int? CanalKds { get; }
    public bool ImprimeArea { get; }
    public bool EsCombo { get; }
    public string? NombreCombo { get; }
    public IReadOnlyList<CondimentoProduccion> Condimentos { get; }

    public bool TieneDestinoKds => !string.IsNullOrWhiteSpace(ObtenerEstacionKds());

    public string NombreSalida
        => EsCombo && !string.IsNullOrWhiteSpace(NombreCombo)
            ? $"{OrdenProduccion.ObtenerIniciales(NombreCombo)}{NombreProducto}"
            : NombreProducto;

    private ItemProduccion(
        string codigoPedido,
        string item,
        string? itemCombo,
        string codigoProducto,
        string nombreProducto,
        decimal cantidad,
        string? codigoArea,
        string? estacionKds,
        int? canalKds,
        bool imprimeArea,
        bool esCombo,
        string? nombreCombo,
        IReadOnlyList<CondimentoProduccion> condimentos)
    {
        CodigoPedido = codigoPedido;
        Item = item;
        ItemCombo = itemCombo;
        CodigoProducto = codigoProducto;
        NombreProducto = nombreProducto;
        Cantidad = cantidad;
        CodigoArea = codigoArea;
        EstacionKds = estacionKds;
        CanalKds = canalKds;
        ImprimeArea = imprimeArea;
        EsCombo = esCombo;
        NombreCombo = nombreCombo;
        Condimentos = condimentos;
    }

    public static ItemProduccion Crear(
        string codigoPedido,
        string item,
        string codigoProducto,
        string nombreProducto,
        decimal cantidad,
        string? codigoArea,
        bool imprimeArea,
        string? itemCombo = null,
        string? estacionKds = null,
        int? canalKds = null,
        bool esCombo = false,
        string? nombreCombo = null,
        IEnumerable<CondimentoProduccion>? condimentos = null)
    {
        if (string.IsNullOrWhiteSpace(codigoPedido))
            throw new DomainException("El código del pedido es obligatorio.", "COCINA_PEDIDO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(item))
            throw new DomainException("El item es obligatorio.", "COCINA_ITEM_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoProducto))
            throw new DomainException("El código del producto es obligatorio.", "COCINA_PRODUCTO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(nombreProducto))
            throw new DomainException("El nombre del producto es obligatorio.", "COCINA_PRODUCTO_NOMBRE_REQUERIDO");

        if (cantidad <= 0)
            throw new DomainException("La cantidad del ítem debe ser mayor a cero.", "COCINA_CANTIDAD_INVALIDA");

        if (esCombo && string.IsNullOrWhiteSpace(nombreCombo))
            throw new DomainException("El nombre del combo es obligatorio para ítems combo.", "COCINA_COMBO_NOMBRE_REQUERIDO");

        return new ItemProduccion(
            codigoPedido.Trim(),
            item.Trim(),
            Limpiar(itemCombo),
            codigoProducto.Trim(),
            nombreProducto.Trim(),
            cantidad,
            Limpiar(codigoArea),
            Limpiar(estacionKds),
            canalKds,
            imprimeArea,
            esCombo,
            Limpiar(nombreCombo),
            condimentos?.ToArray() ?? []);
    }

    public string ObtenerEstacionKds()
        => !string.IsNullOrWhiteSpace(EstacionKds)
            ? EstacionKds!
            : CodigoArea ?? string.Empty;

    public int ObtenerCanalKds()
        => CanalKds ?? 0;

    public string ObtenerIdentificadorKds()
        => string.IsNullOrWhiteSpace(ItemCombo)
            ? Item
            : $"{Item}|{ItemCombo}";

    private static string? Limpiar(string? valor)
        => string.IsNullOrWhiteSpace(valor) ? null : valor.Trim();
}
