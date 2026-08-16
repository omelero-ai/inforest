using Inforest.Domain.Entities.Hardware;

namespace Inforest.Desktop.Caja;

/// <summary>
/// Selector de terminal PinPad por caja.
/// Legacy: frmPagoPinPad.frm, TCAJATERMINAL, TTERMINAL.
/// Regla BR-PERIPH-001.
/// </summary>
public sealed class FrmPagoPinPad : Form
{
    private readonly ListBox _lstTerminales = new() { Dock = DockStyle.Fill, DisplayMember = nameof(TerminalPinPad.Descripcion) };

    public FrmPagoPinPad(IReadOnlyList<TerminalPinPad> terminales)
    {
        Text = "Elige PinPad de Pago";
        Width = 520;
        Height = 360;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        foreach (var terminal in terminales)
            _lstTerminales.Items.Add(terminal);

        if (_lstTerminales.Items.Count > 0)
            _lstTerminales.SelectedIndex = 0;

        var btnAceptar = new Button { Text = "Aceptar", Width = 100 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 100 };
        btnAceptar.Click += (_, _) =>
        {
            if (TerminalSeleccionado is null)
            {
                MessageBox.Show("Seleccione un terminal PinPad.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        };
        btnCancelar.Click += (_, _) =>
        {
            DialogResult = DialogResult.Cancel;
            Close();
        };

        var footer = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 48,
            Padding = new Padding(12),
            FlowDirection = FlowDirection.RightToLeft
        };
        footer.Controls.Add(btnCancelar);
        footer.Controls.Add(btnAceptar);

        Controls.Add(_lstTerminales);
        Controls.Add(footer);
    }

    public TerminalPinPad? TerminalSeleccionado => _lstTerminales.SelectedItem as TerminalPinPad;
}
