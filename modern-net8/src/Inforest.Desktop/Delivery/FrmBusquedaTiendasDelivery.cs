using Inforest.Application.Delivery;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Diálogo modal para seleccionar una tienda/sucursal de un cliente delivery.
/// <para>
/// Legacy: <c>frmBusquedaDelivery.frm</c> Tienda_Click — abre <c>frmBusquedaRapida</c>
/// con datos de <c>vTienda</c>.
/// </para>
/// </summary>
public sealed class FrmBusquedaTiendasDelivery : Form
{
    public TiendaDeliveryItem? TiendaSeleccionada { get; private set; }

    private readonly DataGridView _grid;

    public FrmBusquedaTiendasDelivery(IReadOnlyList<TiendaDeliveryItem> tiendas)
    {
        Text = "Tiendas del Cliente";
        Size = new Size(500, 360);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        KeyPreview = true;

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };
        _grid.Columns.AddRange(
            new DataGridViewTextBoxColumn { Name = "colCodigo", HeaderText = "Código", DataPropertyName = "CodigoTienda", FillWeight = 20 },
            new DataGridViewTextBoxColumn { Name = "colDesc", HeaderText = "Descripción", DataPropertyName = "Descripcion", FillWeight = 40 },
            new DataGridViewTextBoxColumn { Name = "colDir", HeaderText = "Dirección", DataPropertyName = "Direccion", FillWeight = 40 }
        );
        _grid.DataSource = tiendas.ToList();
        _grid.CellDoubleClick += (_, _) => Confirmar();

        var pnlBotones = new Panel { Dock = DockStyle.Bottom, Height = 44 };
        var btnSelec = new Button { Text = "Seleccionar", Left = 100, Top = 8, Width = 110 };
        var btnCancela = new Button { Text = "Cancelar", Left = 220, Top = 8, Width = 90 };
        btnSelec.Click += (_, _) => Confirmar();
        btnCancela.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
        pnlBotones.Controls.AddRange([btnSelec, btnCancela]);

        Controls.AddRange([_grid, pnlBotones]);
        KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Enter) Confirmar();
            if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); }
        };
    }

    private void Confirmar()
    {
        if (_grid.SelectedRows.Count == 0) return;
        var row = _grid.SelectedRows[0];
        TiendaSeleccionada = new TiendaDeliveryItem(
            row.Cells["colCodigo"].Value?.ToString() ?? "",
            row.Cells["colDesc"].Value?.ToString() ?? "",
            row.Cells["colDir"].Value?.ToString());
        DialogResult = DialogResult.OK;
        Close();
    }
}
