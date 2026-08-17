using Inforest.Application.Seguridad;

namespace Inforest.Desktop.POS;

/// <summary>
/// Pantalla de cambio de contraseña del usuario POS.
/// Legacy: frmCambiarContrasenia.frm. BR-006.
/// </summary>
public class FrmCambiarContrasenia : Form
{
    private readonly CambiarPasswordHandler? _handler;
    private readonly TextBox _txtUsuario;
    private readonly TextBox _txtActual;
    private readonly TextBox _txtNueva;
    private readonly TextBox _txtConfirmaNueva;
    private readonly Label _lblInfo;

    public FrmCambiarContrasenia(
        string? loginInicial = null,
        CambiarPasswordHandler? handler = null)
    {
        _handler = handler;

        Text = "Cambiar Contraseña";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(380, 280);

        // ── Labels ─────────────────────────────────────────────────────────────
        var lblUsuario = new Label { Left = 20, Top = 20, Width = 120, Text = "Usuario:" };
        var lblActual = new Label { Left = 20, Top = 60, Width = 120, Text = "Contraseña actual:" };
        var lblNueva = new Label { Left = 20, Top = 100, Width = 120, Text = "Nueva contraseña:" };
        var lblConfirma = new Label { Left = 20, Top = 140, Width = 120, Text = "Confirmar nueva:" };

        // ── Inputs ─────────────────────────────────────────────────────────────
        _txtUsuario = new TextBox
        {
            Left = 150, Top = 17, Width = 200,
            Text = (loginInicial ?? string.Empty).Trim().ToUpperInvariant(),
            ReadOnly = loginInicial is not null
        };
        _txtUsuario.Leave += (_, _) => _txtUsuario.Text = _txtUsuario.Text.Trim().ToUpperInvariant();
        _txtActual = new TextBox { Left = 150, Top = 57, Width = 200, UseSystemPasswordChar = true };
        _txtNueva = new TextBox { Left = 150, Top = 97, Width = 200, UseSystemPasswordChar = true };
        _txtConfirmaNueva = new TextBox { Left = 150, Top = 137, Width = 200, UseSystemPasswordChar = true };

        // ── Info label ─────────────────────────────────────────────────────────
        _lblInfo = new Label
        {
            Left = 20, Top = 180, Width = 340, Height = 36,
            ForeColor = Color.Red,
            AutoSize = false
        };

        // ── Buttons ────────────────────────────────────────────────────────────
        var btnAceptar = new Button
        {
            Left = 130, Top = 225, Width = 100, Text = "Aceptar",
            DialogResult = DialogResult.None
        };
        btnAceptar.Click += async (_, _) => await AceptarAsync();

        var btnCancelar = new Button
        {
            Left = 250, Top = 225, Width = 100, Text = "Cancelar",
            DialogResult = DialogResult.Cancel
        };
        btnCancelar.Click += (_, _) => Close();

        Controls.AddRange([
            lblUsuario, lblActual, lblNueva, lblConfirma,
            _txtUsuario, _txtActual, _txtNueva, _txtConfirmaNueva,
            _lblInfo, btnAceptar, btnCancelar
        ]);

        AcceptButton = btnAceptar;
        CancelButton = btnCancelar;
    }

    private async Task AceptarAsync()
    {
        _lblInfo.Text = string.Empty;

        if (_handler is null)
        {
            _lblInfo.Text = "Servicio no configurado.";
            return;
        }

        var result = await _handler.HandleAsync(new CambiarPasswordCommand(
            _txtUsuario.Text,
            _txtActual.Text,
            _txtNueva.Text,
            _txtConfirmaNueva.Text));

        if (result.EsExitoso)
        {
            MessageBox.Show("Contraseña cambiada exitosamente.", "Seguridad",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
            Close();
        }
        else
        {
            _lblInfo.Text = result.MensajeError ?? "Error desconocido.";
        }
    }
}
