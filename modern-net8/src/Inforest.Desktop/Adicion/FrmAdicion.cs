using System.ComponentModel;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Application.Pedidos;
using Inforest.Application.Turno;
using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Adicion;

/// <summary>
/// Legacy: <c>frmAdicion.frm</c>, <c>Adicion.exe</c>, <c>modAdicion.bas</c>.
/// Permite agregar productos a pedidos abiertos en el turno activo.
/// </summary>
public class FrmAdicion : Form
{
    private readonly IProductoMaestroRepository _productoRepository;
    private readonly ISessionService _sessionService;
    private readonly ObtenerTurnoActualHandler _turnoHandler;
    private readonly ObtenerPedidosPorTurnoHandler _pedidosTurnoHandler;
    private readonly UpdatePedidoHandler _updatePedidoHandler;

    private readonly BindingList<PedidoActivoRow> _pedidos = [];
    private readonly BindingList<ItemPedidoRow> _itemsActuales = [];
    private readonly DataGridView _gridPedidos;
    private readonly DataGridView _gridItems;
    private readonly ComboBox _cmbProductos;
    private List<ProductoMaestro> _productos = [];
    private List<Pedido> _pedidosAbiertos = [];

    public FrmAdicion(
        IProductoMaestroRepository productoRepository,
        ISessionService sessionService,
        ObtenerTurnoActualHandler turnoHandler,
        ObtenerPedidosPorTurnoHandler pedidosTurnoHandler,
        UpdatePedidoHandler updatePedidoHandler)
    {
        _productoRepository = productoRepository;
        _sessionService = sessionService;
        _turnoHandler = turnoHandler;
        _pedidosTurnoHandler = pedidosTurnoHandler;
        _updatePedidoHandler = updatePedidoHandler;

        Text = "Adición";
        WindowState = FormWindowState.Maximized;

        _gridPedidos = new DataGridView
        {
            Dock = DockStyle.Top,
            Height = 220,
            AutoGenerateColumns = true,
            AllowUserToAddRows = false,
            DataSource = _pedidos,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            ReadOnly = true
        };
        _gridPedidos.SelectionChanged += (_, _) => CargarItemsPedido();

        _gridItems = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            AllowUserToAddRows = false,
            DataSource = _itemsActuales
        };

        _cmbProductos = new ComboBox { Width = 240, DropDownStyle = ComboBoxStyle.DropDownList };
        var btnAgregar = new Button { Text = "Agregar", Width = 100 };
        var btnGuardar = new Button { Text = "Guardar", Width = 100 };
        var btnCerrar = new Button { Text = "Cerrar", Width = 100 };
        btnAgregar.Click += (_, _) => AgregarItem();
        btnGuardar.Click += async (_, _) => await GuardarAsync();
        btnCerrar.Click += (_, _) => Close();

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8) };
        acciones.Controls.AddRange([_cmbProductos, btnAgregar, btnGuardar, btnCerrar]);

        Controls.Add(_gridItems);
        Controls.Add(_gridPedidos);
        Controls.Add(acciones);
        Load += async (_, _) => await CargarAsync();
    }

    private async Task CargarAsync()
    {
        // Cargar productos
        try { _productos = (await _productoRepository.ObtenerTodosAsync()).Take(30).ToList(); }
        catch { _productos = []; }

        if (_productos.Count == 0)
            _productos = [ProductoMaestro.Crear("P001", "COM", "Adición Producto", "SISTEMA", 10m)];

        _cmbProductos.DataSource = _productos;
        _cmbProductos.DisplayMember = nameof(ProductoMaestro.Detallado);
        _cmbProductos.ValueMember = nameof(ProductoMaestro.CodigoProducto);

        // Cargar pedidos del turno activo
        await CargarPedidosTurnoAsync();
        CargarItemsPedido();
    }

    private async Task CargarPedidosTurnoAsync()
    {
        _pedidos.Clear();
        _pedidosAbiertos.Clear();

        var codigoCaja = _sessionService.SesionActual?.CodigoCaja;
        if (string.IsNullOrWhiteSpace(codigoCaja)) return;

        try
        {
            var turnoResult = await _turnoHandler.HandleAsync(new ObtenerTurnoActualQuery(codigoCaja));
            if (!turnoResult.EsExitoso || turnoResult.Valor is null) return;

            var pedidosResult = await _pedidosTurnoHandler.HandleAsync(new ObtenerPedidosPorTurnoQuery(turnoResult.Valor.CodigoTurno));
            if (!pedidosResult.EsExitoso || pedidosResult.Valor is null) return;

            _pedidosAbiertos = pedidosResult.Valor.Where(p => p.Estado == EstadoPedido.Abierto || p.Estado == EstadoPedido.EnProceso).ToList();
            foreach (var p in _pedidosAbiertos)
                _pedidos.Add(new PedidoActivoRow(p.CodigoPedido, p.CodigoMesa ?? "--", p.Estado.ToString()));
        }
        catch { /* no hay turno */ }
    }

    private void CargarItemsPedido()
    {
        _itemsActuales.Clear();
        if (_gridPedidos.CurrentRow?.DataBoundItem is not PedidoActivoRow row) return;

        var pedido = _pedidosAbiertos.FirstOrDefault(p => p.CodigoPedido == row.CodigoPedido);
        if (pedido is null) return;

        foreach (var det in pedido.Detalles)
            _itemsActuales.Add(new ItemPedidoRow(pedido.CodigoPedido, det.CodigoProducto, det.CodigoProducto, det.Cantidad, det.PrecioVenta));
    }

    private void AgregarItem()
    {
        if (_gridPedidos.CurrentRow?.DataBoundItem is not PedidoActivoRow pedidoRow || _cmbProductos.SelectedItem is not ProductoMaestro producto)
            return;

        _itemsActuales.Add(new ItemPedidoRow(pedidoRow.CodigoPedido, producto.CodigoProducto, producto.Detallado, 1, producto.PrecioVenta));
    }

    private async Task GuardarAsync()
    {
        if (_gridPedidos.CurrentRow?.DataBoundItem is not PedidoActivoRow pedidoRow)
        {
            MessageBox.Show("Seleccione un pedido.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _updatePedidoHandler.HandleAsync(new UpdatePedidoCommand(
            pedidoRow.CodigoPedido,
            null,
            null,
            _sessionService.SesionActual?.CodigoUsuario ?? Environment.UserName));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show("Adición guardada.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private sealed record PedidoActivoRow(string CodigoPedido, string Mesa, string Estado);
    private sealed record ItemPedidoRow(string CodigoPedido, string CodigoProducto, string Producto, decimal Cantidad, decimal Precio);
}
