namespace Inforest.Domain.Entities.Ventas;

/// <summary>
/// Item de un pedido dentro de una sesión de división.
/// Legacy: fila de la tabla temporal sDetalle en frmDivision.frm.
/// BR-DIV-003, BR-DIV-005.
/// </summary>
public sealed class ItemDivision
{
    /// <summary>Item (tItem) — código de posición de 3 dígitos.</summary>
    public string Item { get; private set; }

    /// <summary>Código del pedido al que pertenece actualmente.</summary>
    public string CodigoPedido { get; private set; }

    /// <summary>
    /// Código del pedido destino temporal. Nulo = permanece en el origen.
    /// Equivale a Estado=' ' (nulo) o Estado='X' (asignado).
    /// </summary>
    public string? CodigoPedidoDestino { get; private set; }

    public string TipoPedido { get; }
    public string CodigoProducto { get; }
    public string NombreProducto { get; set; } = string.Empty;
    public string CodigoGrupo { get; }
    public string CodigoSubGrupo { get; }
    public string Moneda { get; }
    public double PrecioNeto { get; }
    public double PrecioImpuesto1 { get; }
    public double PrecioImpuesto2 { get; }
    public double PrecioImpuesto3 { get; }
    public double PrecioVenta { get; }
    public double Recargo { get; }
    public double Descuento { get; }
    public double PrecioOficial { get; }

    /// <summary>Cantidad actual (ajustable por Disgregar/Compartir).</summary>
    public double Cantidad { get; private set; }

    /// <summary>Impuesto 1 total (proporcional a Cantidad).</summary>
    public double Impuesto1 { get; private set; }
    public double Impuesto2 { get; private set; }
    public double Impuesto3 { get; private set; }

    /// <summary>Venta total (proporcional a Cantidad).</summary>
    public double Venta { get; private set; }

    public string? Observacion { get; }
    public string? Cortesia { get; }
    public bool Imprime { get; }
    public string? EstadoItem { get; }
    public string? Area { get; }
    public bool TienePropiedad { get; set; }

    public ItemDivision(
        string item,
        string codigoPedido,
        string tipoPedido,
        string codigoProducto,
        string codigoGrupo,
        string codigoSubGrupo,
        string moneda,
        double precioNeto,
        double precioImpuesto1,
        double precioImpuesto2,
        double precioImpuesto3,
        double precioVenta,
        double recargo,
        double descuento,
        double precioOficial,
        double cantidad,
        double impuesto1,
        double impuesto2,
        double impuesto3,
        double venta,
        string? observacion,
        string? cortesia,
        bool imprime,
        string? estadoItem,
        string? area)
    {
        Item = item;
        CodigoPedido = codigoPedido;
        TipoPedido = tipoPedido;
        CodigoProducto = codigoProducto;
        CodigoGrupo = codigoGrupo;
        CodigoSubGrupo = codigoSubGrupo;
        Moneda = moneda;
        PrecioNeto = precioNeto;
        PrecioImpuesto1 = precioImpuesto1;
        PrecioImpuesto2 = precioImpuesto2;
        PrecioImpuesto3 = precioImpuesto3;
        PrecioVenta = precioVenta;
        Recargo = recargo;
        Descuento = descuento;
        PrecioOficial = precioOficial;
        Cantidad = cantidad;
        Impuesto1 = impuesto1;
        Impuesto2 = impuesto2;
        Impuesto3 = impuesto3;
        Venta = venta;
        Observacion = observacion;
        Cortesia = cortesia;
        Imprime = imprime;
        EstadoItem = estadoItem;
        Area = area;
    }

    /// <summary>
    /// Asigna el item al pedido destino (Estado='X').
    /// BR-DIV-005.
    /// </summary>
    internal void AsignarDestino(string codigoPedidoDestino)
    {
        CodigoPedidoDestino = codigoPedidoDestino;
    }

    /// <summary>
    /// Revierte el item al pedido origen (Estado=' ').
    /// BR-DIV-005.
    /// </summary>
    internal void RevertirAOrigen(string codigoPedidoOrigen)
    {
        CodigoPedido = codigoPedidoOrigen;
        CodigoPedidoDestino = null;
    }

    /// <summary>
    /// Ajusta la cantidad del item y recalcula impuestos y venta proporcionalmente.
    /// BR-DIV-003.
    /// </summary>
    internal void AjustarCantidad(double nuevaCantidad)
    {
        if (Cantidad == 0) return;
        double factor = nuevaCantidad / Cantidad;
        Impuesto1 *= factor;
        Impuesto2 *= factor;
        Impuesto3 *= factor;
        Venta *= factor;
        Cantidad = nuevaCantidad;
    }

    /// <summary>
    /// Divide impuestos y venta por un divisor (para Compartir).
    /// BR-DIV-004.
    /// </summary>
    internal void DividirProporcionalmente(int divisor)
    {
        if (divisor <= 1) return;
        Cantidad /= divisor;
        Impuesto1 /= divisor;
        Impuesto2 /= divisor;
        Impuesto3 /= divisor;
        Venta /= divisor;
    }

    /// <summary>
    /// Crea una copia del item con el tItem dado y la cantidad especificada,
    /// calculando impuestos proporcionales.
    /// BR-DIV-003.
    /// </summary>
    internal ItemDivision ClonarConCantidad(string nuevoItem, double cantidad)
    {
        double factor = Cantidad > 0 ? cantidad / Cantidad : 0;
        var clon = new ItemDivision(
            item: nuevoItem,
            codigoPedido: CodigoPedido,
            tipoPedido: TipoPedido,
            codigoProducto: CodigoProducto,
            codigoGrupo: CodigoGrupo,
            codigoSubGrupo: CodigoSubGrupo,
            moneda: Moneda,
            precioNeto: PrecioNeto,
            precioImpuesto1: PrecioImpuesto1,
            precioImpuesto2: PrecioImpuesto2,
            precioImpuesto3: PrecioImpuesto3,
            precioVenta: PrecioVenta,
            recargo: Recargo,
            descuento: Descuento,
            precioOficial: PrecioOficial,
            cantidad: cantidad,
            impuesto1: Impuesto1 * factor,
            impuesto2: Impuesto2 * factor,
            impuesto3: Impuesto3 * factor,
            venta: Venta * factor,
            observacion: Observacion,
            cortesia: Cortesia,
            imprime: Imprime,
            estadoItem: EstadoItem,
            area: Area)
        {
            NombreProducto = NombreProducto,
            TienePropiedad = TienePropiedad
        };
        return clon;
    }

    /// <summary>
    /// Crea una copia del item para Compartir, ya dividida y asignada al destino.
    /// BR-DIV-004.
    /// </summary>
    internal ItemDivision ClonarConCantidadProporcional(string nuevoItem, string codigoDestino, int divisor)
    {
        var clon = ClonarConCantidad(nuevoItem, Cantidad);
        clon.AsignarDestino(codigoDestino);
        return clon;
    }
}
