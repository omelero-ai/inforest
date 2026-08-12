using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;

namespace Inforest.Desktop.POS;

/// <summary>
/// Legacy: <c>frmAcceso.frm</c> para el flujo del POS.
/// </summary>
public class FrmLoginPOS : Form
{
    private readonly IAuthService _authService;
    private readonly ObtenerTodasCajasHandler _cajasHandler;
    private readonly TextBox _txtUsuario;
    private readonly TextBox _txtPassword;
    private readonly ComboBox _cmbCaja;
    private readonly Label _lblEstado;

    public FrmLoginPOS(IAuthService authService, ObtenerTodasCajasHandler cajasHandler)
    {
        _authService = authService;
        _cajasHandler = cajasHandler;

        Text = "Login POS";
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(420, 260);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 2, RowCount = 5 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _txtUsuario = new TextBox();
        _txtPassword = new TextBox { UseSystemPasswordChar = true };
        _cmbCaja = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        _lblEstado = new Label { Dock = DockStyle.Fill, ForeColor = Color.DimGray };

        var btnIngresar = new Button { Text = "Ingresar", Width = 100 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 100 };
        btnIngresar.Click += async (_, _) => await IngresarAsync();
        btnCancelar.Click += (_, _) => Close();

        layout.Controls.Add(new Label { Text = "Usuario", AutoSize = true }, 0, 0);
        layout.Controls.Add(_txtUsuario, 1, 0);
        layout.Controls.Add(new Label { Text = "Password", AutoSize = true }, 0, 1);
        layout.Controls.Add(_txtPassword, 1, 1);
        layout.Controls.Add(new Label { Text = "Caja", AutoSize = true }, 0, 2);
        layout.Controls.Add(_cmbCaja, 1, 2);
        layout.Controls.Add(_lblEstado, 0, 3);
        layout.SetColumnSpan(_lblEstado, 2);

        var botones = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        botones.Controls.AddRange([btnCancelar, btnIngresar]);
        layout.Controls.Add(botones, 0, 4);
        layout.SetColumnSpan(botones, 2);

        Controls.Add(layout);
        Load += async (_, _) => await CargarCajasAsync();
    }

    private async Task CargarCajasAsync()
    {
        var result = await _cajasHandler.HandleAsync(new ObtenerTodasCajasQuery());
        _cmbCaja.DataSource = result.Valor?.ToList();
        _cmbCaja.DisplayMember = "tDescripcion";
        _cmbCaja.ValueMember = "tCaja";
    }

    private async Task IngresarAsync()
    {
        var result = await _authService.AutenticarAsync(new AuthRequest(
            _txtUsuario.Text,
            _txtPassword.Text,
            "INFOREST",
            _cmbCaja.SelectedValue?.ToString() ?? "001",
            Environment.MachineName,
            "INFOREST"));

        _lblEstado.Text = result.Exitoso ? "Autenticación correcta." : result.MensajeError ?? "Error de autenticación.";
        if (result.Exitoso)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
