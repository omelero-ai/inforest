using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Selector moderno de clientes para emisión documental.
/// Soporta el flujo legado de frmBusquedaRapida.frm consumido por frmSolicitudBoleta.frm.
/// </summary>
public sealed class FrmSeleccionClienteBoleta : Form
{
    private readonly ObtenerClientesActivosHandler _handler;
    private readonly TextBox _txtFiltro;
    private readonly DataGridView _grid;
    private List<Cliente> _clientes = [];

    public FrmSeleccionClienteBoleta(ObtenerClientesActivosHandler handler)
    {
        _handler = handler;

        Text = "Seleccionar cliente";
        Width = 880;
        Height = 520;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(245, 246, 247);

        _txtFiltro = new TextBox { PlaceholderText = "Buscar por código, identidad o nombre...", Dock = DockStyle.Top, Margin = new Padding(0), Font = new Font("Segoe UI", 10f) };
        _txtFiltro.TextChanged += (_, _) => AplicarFiltro();

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None
        };
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Código", DataPropertyName = nameof(Cliente.CodigoCliente), Width = 90 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Identidad", DataPropertyName = nameof(Cliente.Identidad), Width = 120 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Cliente", DataPropertyName = nameof(Cliente.Empresa), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
        _grid.CellDoubleClick += (_, _) => SeleccionarActual();

        var btnAceptar = new Button { Text = "Aceptar", Width = 100, Height = 34, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        var btnCancelar = new Button { Text = "Cancelar", Width = 100, Height = 34, FlatStyle = FlatStyle.Flat };
        btnAceptar.FlatAppearance.BorderSize = 0;
        btnAceptar.Click += (_, _) => SeleccionarActual();
        btnCancelar.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };

        var acciones = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 52,
            Padding = new Padding(12, 8, 12, 8),
            FlowDirection = FlowDirection.RightToLeft
        };
        acciones.Controls.Add(btnCancelar);
        acciones.Controls.Add(btnAceptar);

        var top = new Panel { Dock = DockStyle.Top, Height = 56, Padding = new Padding(12) };
        top.Controls.Add(_txtFiltro);

        Controls.Add(_grid);
        Controls.Add(acciones);
        Controls.Add(top);

        Load += async (_, _) => await CargarAsync();
    }

    public Cliente? ClienteSeleccionado => _grid.CurrentRow?.DataBoundItem as Cliente;

    private async Task CargarAsync()
    {
        var result = await _handler.HandleAsync(new ObtenerClientesActivosQuery());
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.Cancel;
            Close();
            return;
        }

        _clientes = result.Valor!.OrderBy(static c => c.Empresa).ToList();
        AplicarFiltro();
    }

    private void AplicarFiltro()
    {
        var filtro = _txtFiltro.Text.Trim();
        IEnumerable<Cliente> query = _clientes;

        if (!string.IsNullOrWhiteSpace(filtro))
        {
            query = query.Where(cliente =>
                cliente.CodigoCliente.Contains(filtro, StringComparison.OrdinalIgnoreCase) ||
                cliente.Identidad.Contains(filtro, StringComparison.OrdinalIgnoreCase) ||
                cliente.Empresa.Contains(filtro, StringComparison.OrdinalIgnoreCase));
        }

        _grid.DataSource = query.ToList();
        if (_grid.Rows.Count > 0)
            _grid.Rows[0].Selected = true;
    }

    private void SeleccionarActual()
    {
        if (ClienteSeleccionado is null)
        {
            MessageBox.Show("Seleccione un cliente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        DialogResult = DialogResult.OK;
        Close();
    }
}
