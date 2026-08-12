using System.ComponentModel;

namespace Inforest.Desktop.Caja;

/// <summary>
/// Pantalla de cuadre y cierre de caja.
/// Legacy: frmCaja.frm.
/// </summary>
public class FrmCuadreCaja : Form
{
    private readonly BindingList<CuadreRow> _rows = [];
    private readonly TextBox _txtMontoCierre = new() { Width = 140, Text = "0.00" };
    private readonly Label _lblDiferencia = new() { AutoSize = true };

    public FrmCuadreCaja()
    {
        Text = "Cuadre de Caja";
        Width = 700;
        Height = 460;
        StartPosition = FormStartPosition.CenterParent;

        _rows.Add(new CuadreRow("Efectivo", 120m));
        _rows.Add(new CuadreRow("Tarjeta", 75m));

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            ReadOnly = true,
            DataSource = _rows
        };

        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 72,
            Padding = new Padding(12),
            FlowDirection = FlowDirection.LeftToRight
        };

        var btnCerrarTurno = new Button { Text = "Cerrar turno", Width = 110 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 90 };
        btnCancelar.Click += (_, _) => Close();
        _txtMontoCierre.TextChanged += (_, _) => ActualizarDiferencia();

        panel.Controls.Add(new Label { Text = "Monto cierre:", AutoSize = true, Margin = new Padding(0, 8, 6, 0) });
        panel.Controls.Add(_txtMontoCierre);
        panel.Controls.Add(new Label { Text = "Diferencia:", AutoSize = true, Margin = new Padding(18, 8, 6, 0) });
        panel.Controls.Add(_lblDiferencia);
        panel.Controls.Add(btnCerrarTurno);
        panel.Controls.Add(btnCancelar);

        Controls.Add(grid);
        Controls.Add(panel);

        ActualizarDiferencia();
    }

    private void ActualizarDiferencia()
    {
        _ = decimal.TryParse(_txtMontoCierre.Text, out var cierre);
        var esperado = _rows.Sum(x => x.MontoEsperado);
        _lblDiferencia.Text = $"{cierre - esperado:0.00}";
    }

    private sealed record CuadreRow(string MedioPago, decimal MontoEsperado);
}
