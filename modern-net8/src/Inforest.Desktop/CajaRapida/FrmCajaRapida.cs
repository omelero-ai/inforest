using System.ComponentModel;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.CajaRapida;

/// <summary>
/// Legacy: <c>frmCajaRapida.frm</c>, <c>CajaRapida.exe</c>, <c>modCajaRapida.bas</c>.
/// </summary>
public class FrmCajaRapida : Form
{
    private readonly IProductoMaestroRepository _productoRepository;
    private readonly FlowLayoutPanel _catalogoPanel;
    private readonly BindingList<ItemRapidoRow> _items = [];
    private readonly DataGridView _grid;
    private readonly Label _lblTotal;

    public FrmCajaRapida(IProductoMaestroRepository productoRepository)
    {
        _productoRepository = productoRepository;
        Text = "Caja Rápida";
        WindowState = FormWindowState.Maximized;

        _catalogoPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Left,
            Width = 420,
            AutoScroll = true,
            Padding = new Padding(8)
        };

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            AllowUserToAddRows = false,
            DataSource = _items
        };

        _lblTotal = new Label
        {
            Dock = DockStyle.Bottom,
            Height = 42,
            Padding = new Padding(8),
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            Text = "Total: 0.00"
        };

        var botones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8), FlowDirection = FlowDirection.RightToLeft };
        var btnSalir = new Button { Text = "Salir", Width = 100 };
        var btnLimpiar = new Button { Text = "Limpiar", Width = 100 };
        var btnCobrar = new Button { Text = "Cobrar", Width = 100 };
        btnSalir.Click += (_, _) => Close();
        btnLimpiar.Click += (_, _) => { _items.Clear(); ActualizarTotal(); };
        btnCobrar.Click += (_, _) => Cobrar();
        botones.Controls.AddRange([btnSalir, btnLimpiar, btnCobrar]);

        Controls.Add(_grid);
        Controls.Add(_catalogoPanel);
        Controls.Add(_lblTotal);
        Controls.Add(botones);
        Load += async (_, _) => await CargarProductosAsync();
    }

    private async Task CargarProductosAsync()
    {
        IReadOnlyList<ProductoMaestro> productos;
        try
        {
            productos = await _productoRepository.ObtenerTodosAsync();
        }
        catch
        {
            productos =
            [
                ProductoMaestro.Crear("P001", "BEB", "Gaseosa", "SISTEMA", 8m),
                ProductoMaestro.Crear("P002", "COM", "Hamburguesa", "SISTEMA", 18m),
                ProductoMaestro.Crear("P003", "POS", "Papas fritas", "SISTEMA", 10m)
            ];
        }

        if (productos.Count == 0)
        {
            productos =
            [
                ProductoMaestro.Crear("P001", "BEB", "Gaseosa", "SISTEMA", 8m),
                ProductoMaestro.Crear("P002", "COM", "Hamburguesa", "SISTEMA", 18m)
            ];
        }

        _catalogoPanel.Controls.Clear();
        foreach (var producto in productos.Take(24))
        {
            var boton = new Button
            {
                Width = 120,
                Height = 80,
                Margin = new Padding(6),
                Text = $"{producto.Detallado}\n{producto.PrecioVenta:C}"
            };
            boton.Click += (_, _) => AgregarItem(producto);
            _catalogoPanel.Controls.Add(boton);
        }
    }

    private void AgregarItem(ProductoMaestro producto)
    {
        var existente = _items.FirstOrDefault(i => i.CodigoProducto == producto.CodigoProducto);
        if (existente is null)
        {
            _items.Add(new ItemRapidoRow(producto.CodigoProducto, producto.Detallado, 1, producto.PrecioVenta));
        }
        else
        {
            var index = _items.IndexOf(existente);
            _items[index] = existente with { Cantidad = existente.Cantidad + 1 };
        }

        _grid.Refresh();
        ActualizarTotal();
    }

    private void ActualizarTotal() => _lblTotal.Text = $"Total: {_items.Sum(i => i.Total):C}";

    private void Cobrar()
    {
        using var frmPago = new Form
        {
            Text = "FrmPago",
            Size = new Size(320, 180),
            StartPosition = FormStartPosition.CenterParent
        };
        frmPago.Controls.Add(new Label
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            Text = $"Cobro inmediato\n{_items.Sum(i => i.Total):C}"
        });
        frmPago.ShowDialog(this);
    }

    private sealed record ItemRapidoRow(string CodigoProducto, string Descripcion, int Cantidad, decimal PrecioUnitario)
    {
        public decimal Total => Cantidad * PrecioUnitario;
    }
}
