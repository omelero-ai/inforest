using Inforest.Application.Turno;

namespace Inforest.Desktop.Turno;

public class FrmAperturaTurno : Form
{
    private readonly AbrirTurnoHandler? _abrirTurnoHandler;
    private readonly IDiaContableService? _diaContableService;

    private TextBox txtCaja = null!;
    private TextBox txtUsuario = null!;
    private TextBox txtMontoInicial = null!;
    private Label lblDiaContable = null!;
    private Button btnAbrir = null!;
    private Button btnCancelar = null!;

    public FrmAperturaTurno(AbrirTurnoHandler? abrirTurnoHandler = null, IDiaContableService? diaContableService = null)
    {
        _abrirTurnoHandler = abrirTurnoHandler;
        _diaContableService = diaContableService;
        InitializeComponent();
        Load += OnLoadAsync;
    }

    private void InitializeComponent()
    {
        Text = "Apertura de Turno";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(420, 220);

        var lblCaja = new Label { Left = 20, Top = 20, Width = 100, Text = "Caja:" };
        var lblUsuario = new Label { Left = 20, Top = 55, Width = 100, Text = "Usuario:" };
        var lblMonto = new Label { Left = 20, Top = 90, Width = 100, Text = "Monto Inicial:" };

        txtCaja = new TextBox { Left = 130, Top = 18, Width = 240, Name = "txtCaja" };
        txtUsuario = new TextBox { Left = 130, Top = 53, Width = 240, Name = "txtUsuario" };
        txtMontoInicial = new TextBox { Left = 130, Top = 88, Width = 240, Name = "txtMontoInicial", Text = "0.00" };
        lblDiaContable = new Label { Left = 20, Top = 130, Width = 350, Text = "Día contable: cargando..." };
        btnAbrir = new Button { Left = 130, Top = 165, Width = 110, Text = "Abrir", Name = "btnAbrir" };
        btnCancelar = new Button { Left = 260, Top = 165, Width = 110, Text = "Cancelar", Name = "btnCancelar" };

        btnCancelar.Click += (_, _) => Close();
        btnAbrir.Click += async (_, _) => await AbrirAsync();

        Controls.AddRange([lblCaja, txtCaja, lblUsuario, txtUsuario, lblMonto, txtMontoInicial, lblDiaContable, btnAbrir, btnCancelar]);
    }

    private async void OnLoadAsync(object? sender, EventArgs e)
    {
        if (_diaContableService is null)
        {
            lblDiaContable.Text = "Día contable: servicio no configurado";
            return;
        }

        try
        {
            var fecha = await _diaContableService.ObtenerDiaContableAsync(true, "05:00", txtUsuario.Text.Trim(), CancellationToken.None);
            lblDiaContable.Text = $"Día contable: {fecha:dd/MM/yyyy}";
        }
        catch
        {
            lblDiaContable.Text = "Día contable: no disponible";
        }
    }

    private async Task AbrirAsync()
    {
        if (_abrirTurnoHandler is null)
        {
            MessageBox.Show("Handler de apertura no configurado.", "Turno", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!decimal.TryParse(txtMontoInicial.Text, out var montoInicial))
        {
            MessageBox.Show("Ingrese un monto inicial válido.", "Turno", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _abrirTurnoHandler.HandleAsync(new AbrirTurnoCommand(
            txtCaja.Text.Trim(),
            txtUsuario.Text.Trim(),
            montoInicial));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Turno", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        DialogResult = DialogResult.OK;
        Close();
    }
}
