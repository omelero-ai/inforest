namespace Inforest.Domain.Entities.Ventas;

/// <summary>
/// Sesión en memoria para la división de un pedido.
/// Legacy: frmDivision.frm — usa tablas temporales SQL; aquí se replantea con colecciones en memoria.
/// Reglas: BR-DIV-001, BR-DIV-002, BR-DIV-003, BR-DIV-004, BR-DIV-005.
/// </summary>
public sealed class SesionDivision
{
    private readonly List<ItemDivision> _items = [];
    private readonly List<PedidoDestino> _pedidosDestino = [];

    /// <summary>Código del pedido origen que se está dividiendo.</summary>
    public string CodigoPedidoOrigen { get; }

    /// <summary>Tipo de pedido del origen (heredado por los nuevos pedidos).</summary>
    public string TipoPedido { get; }

    /// <summary>Items del pedido origen en la sesión (incluye disgregados).</summary>
    public IReadOnlyList<ItemDivision> Items => _items.AsReadOnly();

    /// <summary>Pedidos destino temporales a los que se pueden mover items.</summary>
    public IReadOnlyList<PedidoDestino> PedidosDestino => _pedidosDestino.AsReadOnly();

    private SesionDivision(string codigoPedidoOrigen, string tipoPedido)
    {
        CodigoPedidoOrigen = codigoPedidoOrigen;
        TipoPedido = tipoPedido;
    }

    /// <summary>
    /// Crea una sesión de división cargando los items del pedido origen.
    /// </summary>
    public static SesionDivision Crear(string codigoPedidoOrigen, string tipoPedido, IEnumerable<ItemDivision> items)
    {
        var sesion = new SesionDivision(codigoPedidoOrigen, tipoPedido);
        sesion._items.AddRange(items);
        return sesion;
    }

    // ─── Gestión de pedidos destino ────────────────────────────────────────────

    /// <summary>
    /// Agrega un pedido destino temporal con código secuencial ("001", "002", …).
    /// BR-DIV-005: pedido destino temporal hasta confirmar.
    /// </summary>
    public PedidoDestino AgregarPedidoDestino(string correlativo, string observacion)
    {
        var pedido = new PedidoDestino(
            CodigoTemporal: GenerarCodigoTemporal(),
            Correlativo: correlativo,
            Observacion: observacion);
        _pedidosDestino.Add(pedido);
        return pedido;
    }

    /// <summary>
    /// Elimina el último pedido destino si no tiene items asignados.
    /// </summary>
    public bool QuitarUltimoPedidoDestino()
    {
        if (_pedidosDestino.Count == 0) return false;
        var ultimo = _pedidosDestino[^1];
        if (_items.Any(i => i.CodigoPedidoDestino == ultimo.CodigoTemporal))
            return false; // BR-DIV-005: no eliminar si tiene items
        _pedidosDestino.RemoveAt(_pedidosDestino.Count - 1);
        return true;
    }

    /// <summary>Establece N pedidos destino, descartando los existentes (si sin items).</summary>
    public bool EstablecerCantidadPedidos(int cantidad, string correlativo, string prefijioObservacion)
    {
        // Validar que no haya pedidos con items que serían eliminados
        if (cantidad < _pedidosDestino.Count)
        {
            var aBorrar = _pedidosDestino.Skip(cantidad).ToList();
            if (aBorrar.Any(p => _items.Any(i => i.CodigoPedidoDestino == p.CodigoTemporal)))
                return false;
        }

        _pedidosDestino.Clear();
        for (int i = 1; i <= cantidad; i++)
        {
            var codigo = PadTres(i);
            _pedidosDestino.Add(new PedidoDestino(
                CodigoTemporal: codigo,
                Correlativo: correlativo,
                Observacion: $"{prefijioObservacion} - {codigo}"));
        }
        // Reset items destino si se reconfiguraron pedidos
        foreach (var item in _items.Where(i => i.CodigoPedidoDestino is not null).ToList())
            item.RevertirAOrigen(CodigoPedidoOrigen);
        return true;
    }

    // ─── Movimiento de items ────────────────────────────────────────────────────

    /// <summary>
    /// Mueve un item del origen al pedido destino seleccionado.
    /// BR-DIV-005: Estado ' '→'X'.
    /// </summary>
    public bool MoverItemADestino(string tItem, string codigoPedidoDestino)
    {
        var item = _items.FirstOrDefault(i => i.Item == tItem && i.CodigoPedidoDestino is null);
        if (item is null) return false;
        var destino = _pedidosDestino.FirstOrDefault(p => p.CodigoTemporal == codigoPedidoDestino);
        if (destino is null) return false;
        item.AsignarDestino(codigoPedidoDestino);
        return true;
    }

    /// <summary>
    /// Regresa un item del destino al origen.
    /// BR-DIV-005: Estado 'X'→' '.
    /// </summary>
    public bool RevertirItemAOrigen(string tItem, string codigoPedidoDestino)
    {
        var item = _items.FirstOrDefault(i => i.Item == tItem && i.CodigoPedidoDestino == codigoPedidoDestino);
        if (item is null) return false;
        item.RevertirAOrigen(CodigoPedidoOrigen);
        return true;
    }

    /// <summary>Mueve todos los items sin destino al pedido destino seleccionado.</summary>
    public int MoverTodosADestino(string codigoPedidoDestino)
    {
        var destino = _pedidosDestino.FirstOrDefault(p => p.CodigoTemporal == codigoPedidoDestino);
        if (destino is null) return 0;
        var sinDestino = _items.Where(i => i.CodigoPedidoDestino is null).ToList();
        foreach (var item in sinDestino)
            item.AsignarDestino(codigoPedidoDestino);
        return sinDestino.Count;
    }

    /// <summary>Regresa todos los items del pedido destino al origen.</summary>
    public int RevertirTodosAOrigen(string codigoPedidoDestino)
    {
        var conDestino = _items.Where(i => i.CodigoPedidoDestino == codigoPedidoDestino).ToList();
        foreach (var item in conDestino)
            item.RevertirAOrigen(CodigoPedidoOrigen);
        return conDestino.Count;
    }

    // ─── Disgregar ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Disgrega un item con cantidad > 1 en varios items individuales (cantidad=1 c/u)
    /// o en dos bloques con cantidad especificada.
    /// BR-DIV-003: impuestos y venta se ajustan proporcionalmente.
    /// </summary>
    public bool Disgregar(string tItem, double? cantidadPrimera = null)
    {
        var item = _items.FirstOrDefault(i => i.Item == tItem && i.CodigoPedidoDestino is null);
        if (item is null || item.Cantidad <= 1) return false;

        if (cantidadPrimera.HasValue)
        {
            // División en dos bloques
            if (cantidadPrimera.Value <= 0 || cantidadPrimera.Value >= item.Cantidad) return false;
            var segunda = item.ClonarConCantidad(GenerarSiguienteItem(), item.Cantidad - cantidadPrimera.Value);
            item.AjustarCantidad(cantidadPrimera.Value);
            _items.Add(segunda);
        }
        else
        {
            // Disgregación total en unidades individuales
            double cantidadTotal = item.Cantidad;
            item.AjustarCantidad(1);
            for (int i = 1; i < (int)cantidadTotal; i++)
            {
                var nuevo = item.ClonarConCantidad(GenerarSiguienteItem(), 1);
                _items.Add(nuevo);
            }
        }
        return true;
    }

    // ─── Compartir ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Comparte todos los items del origen equitativamente entre el pedido origen
    /// y todos los pedidos destino.
    /// BR-DIV-004: nCantidad e impuestos se dividen entre (nPedidosDestino + 1).
    /// </summary>
    public bool Compartir()
    {
        if (_pedidosDestino.Count == 0) return false;
        int total = _pedidosDestino.Count + 1;
        var itemsOrigen = _items.Where(i => i.CodigoPedidoDestino is null).ToList();

        foreach (var origen in itemsOrigen)
        {
            origen.DividirProporcionalmente(total);
            foreach (var destino in _pedidosDestino)
            {
                var copia = origen.ClonarConCantidadProporcional(GenerarSiguienteItem(), destino.CodigoTemporal, total);
                _items.Add(copia);
            }
        }
        return true;
    }

    /// <summary>Asigna el código definitivo de MPEDIDO a un pedido destino.</summary>
    public void AsignarCodigoDefinitivoPedido(string codigoTemporal, string codigoDefinitivo)
    {
        var pedido = _pedidosDestino.FirstOrDefault(p => p.CodigoTemporal == codigoTemporal);
        if (pedido is not null)
            pedido.CodigoDefinitivo = codigoDefinitivo;
    }

    // ─── Consultas ──────────────────────────────────────────────────────────────

    /// <summary>Items que permanecen en el pedido origen (sin destino asignado).</summary>
    public IEnumerable<ItemDivision> ItemsOrigen =>
        _items.Where(i => i.CodigoPedidoDestino is null);

    /// <summary>Items asignados a un pedido destino específico.</summary>
    public IEnumerable<ItemDivision> ItemsDestino(string codigoPedidoDestino) =>
        _items.Where(i => i.CodigoPedidoDestino == codigoPedidoDestino);

    /// <summary>Total venta de los items origen.</summary>
    public decimal MontoOrigen => (decimal)_items
        .Where(i => i.CodigoPedidoDestino is null)
        .Sum(i => i.Venta);

    // ─── Helpers privados ──────────────────────────────────────────────────────

    private string GenerarCodigoTemporal()
    {
        int n = _pedidosDestino.Count + 1;
        return PadTres(n);
    }

    private string GenerarSiguienteItem()
    {
        int max = _items
            .Select(i => int.TryParse(i.Item, out var v) ? v : 0)
            .DefaultIfEmpty(0)
            .Max();
        return PadTres(max + 1);
    }

    private static string PadTres(int n) =>
        n.ToString().PadLeft(3, '0');
}

/// <summary>Pedido destino temporal creado durante la división.</summary>
public sealed record PedidoDestino(
    string CodigoTemporal,
    string Correlativo,
    string Observacion)
{
    /// <summary>Código definitivo asignado en la confirmación (nulo hasta confirmar).</summary>
    public string? CodigoDefinitivo { get; internal set; }
}
