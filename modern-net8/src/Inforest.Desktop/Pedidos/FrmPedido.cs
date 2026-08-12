using Inforest.Application.Pedidos;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Pedidos;

public class FrmPedido : Form
{
    private readonly CreatePedidoHandler? _createPedidoHandler;

    private DataGridView dgvProductos = null!;
    private DataGridView dgvDetalle = null!;
    private TextBox txtMesa = null!;
    private TextBox txtMozo = null!;
    private TextBox txtObservacion = null!;
    private Label lblTurno = null!;
    private Label lblCaja = null!;

    public FrmPedido(CreatePedidoHandler? createPedidoHandler = null)
    {
        _createPedidoHandler = createPedidoHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Pedido";
        WindowState = FormWindowState.Maximized;

        var topPanel = new Panel { Dock = DockStyle.Top, Height = 70 };
        lblCaja = new Label { Left = 20, Top = 15, Width = 180, Text = "Caja:" };
        lblTurno = new Label { Left = 220, Top = 15, Width = 180, Text = "Turno:" };
        var lblMesa = new Label { Left = 20, Top = 42, Width = 40, Text = "Mesa:" };
        txtMesa = new TextBox { Left = 65, Top = 39, Width = 80, Name = "txtMesa" };
        var lblMozo = new Label { Left = 160, Top = 42, Width = 45, Text = "Mozo:" };
        txtMozo = new TextBox { Left = 210, Top = 39, Width = 100, Name = "txtMozo" };
        var lblObs = new Label { Left = 325, Top = 42, Width = 85, Text = "Observación:" };
        txtObservacion = new TextBox { Left = 415, Top = 39, Width = 260, Name = "txtObservacion" };
        topPanel.Controls.AddRange([lblCaja, lblTurno, lblMesa, txtMesa, lblMozo, txtMozo, lblObs, txtObservacion]);

        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 450 };

        dgvProductos = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = true
        };
        dgvProductos.DataSource = new[]
        {
            new { CodigoProducto = "P001", Descripcion = "Producto 1", Precio = 10.00m, Area = "COC" },
            new { CodigoProducto = "P002", Descripcion = "Producto 2", Precio = 15.00m, Area = "BAR" }
        };

        dgvDetalle = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = true
        };
        dgvDetalle.DataSource = new BindingSource { DataSource = new List<DetalleGridRow>() };

        split.Panel1.Controls.Add(dgvProductos);
        split.Panel2.Controls.Add(dgvDetalle);

        var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };
        var btnAgregar = new Button { Left = 20, Top = 10, Width = 110, Text = "Agregar", Name = "btnAgregar" };
        var btnEliminar = new Button { Left = 140, Top = 10, Width = 110, Text = "Eliminar", Name = "btnEliminar" };
        var btnEnviarCocina = new Button { Left = 260, Top = 10, Width = 130, Text = "Enviar Cocina", Name = "btnEnviarCocina" };
        var btnCerrar = new Button { Left = 400, Top = 10, Width = 110, Text = "Cerrar", Name = "btnCerrar" };
        bottomPanel.Controls.AddRange([btnAgregar, btnEliminar, btnEnviarCocina, btnCerrar]);

        btnAgregar.Click += (_, _) => AgregarProducto();
        btnEliminar.Click += (_, _) => EliminarProducto();
        btnEnviarCocina.Click += async (_, _) => await EnviarPedidoAsync();
        btnCerrar.Click += (_, _) => Close();

        Controls.Add(split);
        Controls.Add(bottomPanel);
        Controls.Add(topPanel);
    }

    private void AgregarProducto()
    {
        if (dgvProductos.CurrentRow?.DataBoundItem is null)
            return;

        var codigo = dgvProductos.CurrentRow.Cells["CodigoProducto"].Value?.ToString() ?? string.Empty;
        var descripcion = dgvProductos.CurrentRow.Cells["Descripcion"].Value?.ToString() ?? string.Empty;
        var precio = Convert.ToDecimal(dgvProductos.CurrentRow.Cells["Precio"].Value ?? 0m);
        var area = dgvProductos.CurrentRow.Cells["Area"].Value?.ToString();

        var source = (BindingSource)dgvDetalle.DataSource!;
        var rows = (List<DetalleGridRow>)source.DataSource!;
        rows.Add(new DetalleGridRow(codigo, descripcion, 1m, precio, area));
        source.ResetBindings(false);
    }

    private void EliminarProducto()
    {
        if (dgvDetalle.CurrentRow?.DataBoundItem is not DetalleGridRow row)
            return;

        var source = (BindingSource)dgvDetalle.DataSource!;
        var rows = (List<DetalleGridRow>)source.DataSource!;
        rows.Remove(row);
        source.ResetBindings(false);
    }

    private async Task EnviarPedidoAsync()
    {
        if (_createPedidoHandler is null)
        {
            MessageBox.Show("Handler de pedidos no configurado.", "Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var source = (BindingSource)dgvDetalle.DataSource!;
        var rows = ((List<DetalleGridRow>)source.DataSource!).ToList();
        var command = new CreatePedidoCommand(
            CodigoCaja: string.Empty,
            CodigoTurno: string.Empty,
            CodigoMesa: txtMesa.Text.Trim(),
            CodigoMozo: txtMozo.Text.Trim(),
            CodigoUsuario: Environment.UserName,
            Canal: CanalVenta.Local,
            NumeroAdultos: 1,
            NumeroNinos: 0,
            Observacion: txtObservacion.Text.Trim(),
            Items: rows.Select(r => new CreateDetalleItem(r.CodigoProducto, string.Empty, string.Empty, r.Cantidad, r.PrecioUnitario, r.Area, null)).ToList());

        var result = await _createPedidoHandler.HandleAsync(command);
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show($"Pedido {result.Valor?.CodigoPedido} registrado.", "Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private sealed record DetalleGridRow(string CodigoProducto, string Descripcion, decimal Cantidad, decimal PrecioUnitario, string? Area);
}
