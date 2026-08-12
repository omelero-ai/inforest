using System.ComponentModel;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Pantalla simplificada de emisión de documentos de venta.
/// Legacy: frmVenta.frm.
/// </summary>
public class FrmVenta : Form
{
    private readonly BindingList<VentaItemRow> _items = [];
    private readonly Label _lblTotal = new() { AutoSize = true };
    private readonly Label _lblNeto = new() { AutoSize = true };
    private readonly Label _lblImpuesto = new() { AutoSize = true };
    private readonly NumericUpDown _nudPropina = new() { DecimalPlaces = 2, Maximum = 999999, Width = 140 };
    private readonly NumericUpDown _nudDescuento = new() { DecimalPlaces = 2, Maximum = 999999, Width = 140 };

    public FrmVenta()
    {
        Text = "Venta / Emisión de Documento";
        Width = 1000;
        Height = 620;
        StartPosition = FormStartPosition.CenterScreen;

        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            SplitterDistance = 620
        };

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            ReadOnly = true,
            DataSource = _items,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        split.Panel1.Controls.Add(grid);

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 10,
            Padding = new Padding(12),
            AutoSize = true
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));

        var cboTipoDocumento = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
        cboTipoDocumento.Items.AddRange(["01 - Boleta", "03 - Factura"]);
        cboTipoDocumento.SelectedIndex = 0;

        var txtCliente = new TextBox { Width = 200 };
        _nudPropina.ValueChanged += (_, _) => ActualizarTotales();
        _nudDescuento.ValueChanged += (_, _) => ActualizarTotales();

        var btnEmitir = new Button { Text = "Emitir", Width = 100 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 100 };
        btnCancelar.Click += (_, _) => Close();

        panel.Controls.Add(new Label { Text = "Tipo documento:", AutoSize = true }, 0, 0);
        panel.Controls.Add(cboTipoDocumento, 1, 0);
        panel.Controls.Add(new Label { Text = "Cliente:", AutoSize = true }, 0, 1);
        panel.Controls.Add(txtCliente, 1, 1);
        panel.Controls.Add(new Label { Text = "Propina:", AutoSize = true }, 0, 2);
        panel.Controls.Add(_nudPropina, 1, 2);
        panel.Controls.Add(new Label { Text = "Descuento:", AutoSize = true }, 0, 3);
        panel.Controls.Add(_nudDescuento, 1, 3);
        panel.Controls.Add(new Label { Text = "Neto:", AutoSize = true }, 0, 4);
        panel.Controls.Add(_lblNeto, 1, 4);
        panel.Controls.Add(new Label { Text = "Impuestos:", AutoSize = true }, 0, 5);
        panel.Controls.Add(_lblImpuesto, 1, 5);
        panel.Controls.Add(new Label { Text = "Total:", AutoSize = true }, 0, 6);
        panel.Controls.Add(_lblTotal, 1, 6);

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
        acciones.Controls.Add(btnEmitir);
        acciones.Controls.Add(btnCancelar);
        panel.Controls.Add(acciones, 1, 7);

        split.Panel2.Controls.Add(panel);
        Controls.Add(split);

        _items.Add(new VentaItemRow("001", "PRD001", "Producto demo", 1, 10m, 11.80m, 1.80m));
        _items.Add(new VentaItemRow("002", "PRD002", "Producto demo 2", 2, 5m, 5.90m, 1.80m));
        ActualizarTotales();
    }

    private void ActualizarTotales()
    {
        var neto = _items.Sum(x => x.Cantidad * x.PrecioNeto);
        var impuesto = _items.Sum(x => x.ImpuestoTotal);
        var total = neto + impuesto + _nudPropina.Value - _nudDescuento.Value;

        _lblNeto.Text = $"{neto:0.00}";
        _lblImpuesto.Text = $"{impuesto:0.00}";
        _lblTotal.Text = $"{total:0.00}";
    }

    private sealed record VentaItemRow(
        string Item,
        string CodigoProducto,
        string Descripcion,
        decimal Cantidad,
        decimal PrecioNeto,
        decimal PrecioVenta,
        decimal ImpuestoTotal);
}
