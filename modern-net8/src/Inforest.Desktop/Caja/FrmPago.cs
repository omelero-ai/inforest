using System.ComponentModel;

namespace Inforest.Desktop.Caja;

/// <summary>
/// Pantalla simplificada de cobranza.
/// Legacy: frmPago.frm.
/// </summary>
public class FrmPago : Form
{
    private readonly BindingList<PagoGridRow> _pagos = [];
    private readonly Label _lblVuelto = new() { AutoSize = true };
    private readonly Label _lblTotal = new() { AutoSize = true, Font = new Font("Segoe UI", 12, FontStyle.Bold) };
    private readonly NumericUpDown _txtMontoPago = new() { DecimalPlaces = 2, Maximum = 999999, Width = 160, Value = 23.60m };

    public FrmPago()
    {
        Text = "Cobro de Documento";
        Width = 760;
        Height = 520;
        StartPosition = FormStartPosition.CenterParent;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 150,
            Padding = new Padding(12),
            ColumnCount = 2
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));

        var cboMedios = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        cboMedios.Items.AddRange(["Efectivo", "Tarjeta", "Cheque", "Vale", "Otro"]);
        cboMedios.SelectedIndex = 0;

        _lblTotal.Text = "23.60";
        _txtMontoPago.ValueChanged += (_, _) => ActualizarVuelto();

        panel.Controls.Add(new Label { Text = "Total a pagar:", AutoSize = true }, 0, 0);
        panel.Controls.Add(_lblTotal, 1, 0);
        panel.Controls.Add(new Label { Text = "Medio de pago:", AutoSize = true }, 0, 1);
        panel.Controls.Add(cboMedios, 1, 1);
        panel.Controls.Add(new Label { Text = "Monto pago:", AutoSize = true }, 0, 2);
        panel.Controls.Add(_txtMontoPago, 1, 2);
        panel.Controls.Add(new Label { Text = "Vuelto:", AutoSize = true }, 0, 3);
        panel.Controls.Add(_lblVuelto, 1, 3);

        var btnAgregar = new Button { Text = "Agregar", Width = 90 };
        var btnCobrar = new Button { Text = "Cobrar", Width = 90 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 90 };
        btnCancelar.Click += (_, _) => Close();
        btnAgregar.Click += (_, _) =>
        {
            _pagos.Add(new PagoGridRow(cboMedios.Text, _txtMontoPago.Value, Math.Max(0, _txtMontoPago.Value - TotalDocumento)));
        };

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(12, 0, 12, 0) };
        acciones.Controls.Add(btnAgregar);
        acciones.Controls.Add(btnCobrar);
        acciones.Controls.Add(btnCancelar);

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            ReadOnly = true,
            DataSource = _pagos,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        Controls.Add(grid);
        Controls.Add(acciones);
        Controls.Add(panel);

        ActualizarVuelto();
    }

    private decimal TotalDocumento => 23.60m;

    private void ActualizarVuelto()
    {
        var vuelto = Math.Max(0, _txtMontoPago.Value - TotalDocumento);
        _lblVuelto.Text = $"{vuelto:0.00}";
    }

    private sealed record PagoGridRow(string MedioPago, decimal Monto, decimal Vuelto);
}
