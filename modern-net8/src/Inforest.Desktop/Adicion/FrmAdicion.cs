using System.ComponentModel;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Adicion;

/// <summary>
/// Legacy: <c>frmAdicion.frm</c>, <c>Adicion.exe</c>, <c>modAdicion.bas</c>.
/// </summary>
public class FrmAdicion : Form
{
    private readonly IProductoMaestroRepository _productoRepository;
    private readonly BindingList<PedidoActivoRow> _pedidos = [];
    private readonly BindingList<ItemPedidoRow> _itemsActuales = [];
    private readonly DataGridView _gridPedidos;
    private readonly DataGridView _gridItems;
    private readonly ComboBox _cmbProductos;
    private List<ProductoMaestro> _productos = [];

    public FrmAdicion(IProductoMaestroRepository productoRepository)
    {
        _productoRepository = productoRepository;
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
        var btnCerrar = new Button { Text = "Cerrar", Width = 100 };
        btnAgregar.Click += (_, _) => AgregarItem();
        btnCerrar.Click += (_, _) => Close();

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8) };
        acciones.Controls.AddRange([_cmbProductos, btnAgregar, btnCerrar]);

        Controls.Add(_gridItems);
        Controls.Add(_gridPedidos);
        Controls.Add(acciones);
        Load += async (_, _) => await CargarAsync();
    }

    private async Task CargarAsync()
    {
        _pedidos.Clear();
        _pedidos.Add(new PedidoActivoRow("PED-001", "Mesa 01", "Abierto"));
        _pedidos.Add(new PedidoActivoRow("PED-002", "Mesa 05", "Abierto"));

        try
        {
            _productos = (await _productoRepository.ObtenerTodosAsync()).Take(20).ToList();
        }
        catch
        {
            _productos =
            [
                ProductoMaestro.Crear("P001", "COM", "Adición Pizza", "SISTEMA", 25m),
                ProductoMaestro.Crear("P002", "BEB", "Adición Bebida", "SISTEMA", 7m)
            ];
        }

        if (_productos.Count == 0)
        {
            _productos =
            [
                ProductoMaestro.Crear("P001", "COM", "Adición Pizza", "SISTEMA", 25m),
                ProductoMaestro.Crear("P002", "BEB", "Adición Bebida", "SISTEMA", 7m)
            ];
        }

        _cmbProductos.DataSource = _productos;
        _cmbProductos.DisplayMember = nameof(ProductoMaestro.Detallado);
        _cmbProductos.ValueMember = nameof(ProductoMaestro.CodigoProducto);
        CargarItemsPedido();
    }

    private void CargarItemsPedido()
    {
        _itemsActuales.Clear();
        if (_gridPedidos.CurrentRow?.DataBoundItem is not PedidoActivoRow pedido)
            return;

        _itemsActuales.Add(new ItemPedidoRow(pedido.CodigoPedido, "PBASE", "Pedido Base", 1, 18m));
    }

    private void AgregarItem()
    {
        if (_gridPedidos.CurrentRow?.DataBoundItem is not PedidoActivoRow pedido || _cmbProductos.SelectedItem is not ProductoMaestro producto)
            return;

        _itemsActuales.Add(new ItemPedidoRow(pedido.CodigoPedido, producto.CodigoProducto, producto.Detallado, 1, producto.PrecioVenta));
    }

    private sealed record PedidoActivoRow(string CodigoPedido, string Mesa, string Estado);
    private sealed record ItemPedidoRow(string CodigoPedido, string CodigoProducto, string Producto, int Cantidad, decimal Precio);
}
