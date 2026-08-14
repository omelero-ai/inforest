using Inforest.Application.Turno;

namespace Inforest.Desktop.Turno;

/// <summary>
/// Formulario de gestión del Día Contable.
///
/// Legacy: frmDiaContable.frm
/// Modo "Apertura" → permite aperturar un nuevo día contable (si no existe ninguno activo).
/// Modo "Cierre"   → permite cerrar el día contable activo.
///
/// Reglas: BR-DC-001, BR-DC-002, BR-DC-003, BR-DC-004.
/// Legacy: clsDiaContable.cls — usp_GenObtieneDiaContable, usp_GenInsertarDiaContable, UPDATE TDIACONTABLE.
/// </summary>
public class FrmDiaContable : Form
{
    public enum Modo { Apertura, Cierre }

    private readonly AperturarDiaContableHandler? _aperturarHandler;
    private readonly CerrarDiaContableHandler? _cerrarHandler;
    private readonly ObtenerDiaContableHandler? _obtenerHandler;
    private readonly string _usuario;
    private readonly Modo _modo;

    private DateTimePicker dtpDiaContable = null!;
    private Label lblDiaContableActual = null!;
    private Label lblMensaje = null!;
    private Button btnApertura = null!;
    private Button btnCerrar = null!;
    private Button btnSalir = null!;

    // Legacy: lIniciaPorDiaContable — indica si el POS puede iniciar venta
    public bool IniciaPorDiaContable { get; private set; }

    // Legacy: lDiaContableAperturado
    public bool DiaContableAperturado { get; private set; }

    public FrmDiaContable(
        Modo modo,
        string usuario,
        AperturarDiaContableHandler? aperturarHandler = null,
        CerrarDiaContableHandler? cerrarHandler = null,
        ObtenerDiaContableHandler? obtenerHandler = null)
    {
        _modo = modo;
        _usuario = usuario;
        _aperturarHandler = aperturarHandler;
        _cerrarHandler = cerrarHandler;
        _obtenerHandler = obtenerHandler;
        InitializeComponent();
        Load += async (_, _) => await OnLoadAsync();
    }

    private void InitializeComponent()
    {
        Text = "Mensaje día Contable";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(360, 180);

        var grp = new GroupBox { Left = 5, Top = 0, Width = 345, Height = 170 };

        lblDiaContableActual = new Label
        {
            Left = 10, Top = 30, Width = 185, Height = 22,
            Text = "Día Contable Actual :",
            Font = new Font("MS Sans Serif", 12, FontStyle.Bold),
            ForeColor = Color.Blue,
            TextAlign = ContentAlignment.MiddleRight
        };

        dtpDiaContable = new DateTimePicker
        {
            Left = 205, Top = 28, Width = 130,
            Format = DateTimePickerFormat.Short,
            Enabled = false
        };

        lblMensaje = new Label
        {
            Left = 5, Top = 140, Width = 330, Height = 22,
            ForeColor = Color.FromArgb(0, 0, 192),
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Bold),
            BorderStyle = BorderStyle.FixedSingle,
            TextAlign = ContentAlignment.MiddleCenter
        };

        btnApertura = new Button
        {
            Left = 10, Top = 80, Width = 100, Height = 50,
            Text = "Apertura",
            Visible = false,
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Bold)
        };
        btnApertura.Click += async (_, _) => await OnAperturaClickAsync();

        btnCerrar = new Button
        {
            Left = 10, Top = 80, Width = 100, Height = 50,
            Text = "Cierre",
            Visible = false,
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Bold)
        };
        btnCerrar.Click += async (_, _) => await OnCerrarClickAsync();

        btnSalir = new Button
        {
            Left = 240, Top = 80, Width = 100, Height = 50,
            Text = "Salir",
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Bold)
        };
        btnSalir.Click += OnSalirClick;

        grp.Controls.AddRange([lblDiaContableActual, dtpDiaContable, lblMensaje, btnApertura, btnCerrar, btnSalir]);
        Controls.Add(grp);
    }

    private async Task OnLoadAsync()
    {
        DiaContableAperturado = false;

        DateTime fechaDiaContable;
        if (_obtenerHandler is not null)
        {
            var result = await _obtenerHandler.HandleAsync(new ObtenerDiaContableQuery(_usuario));
            fechaDiaContable = result.EsExitoso ? result.Valor!.Fecha : DateTime.Today;
        }
        else
        {
            fechaDiaContable = DateTime.Today;
        }

        dtpDiaContable.Value = fechaDiaContable;

        if (_modo == Modo.Apertura)
        {
            lblDiaContableActual.Visible = true;

            // Legacy: if Format(fdiacontable,"yyyyMMdd") = "19000101" → primer arranque
            var esNuevo = fechaDiaContable.Date == new DateTime(1900, 1, 1).Date;
            if (esNuevo)
            {
                // Sin día contable previo → permitir apertura
                lblDiaContableActual.Text = "Día Contable a Aperturar :";
                lblMensaje.Text = "Proceso de apertura del día contable";
                dtpDiaContable.Value = DateTime.Today;
                dtpDiaContable.Enabled = true;
                btnApertura.Visible = true;
            }
            else
            {
                // BR-DC-002: ya existe día contable activo
                dtpDiaContable.Enabled = false;
                lblDiaContableActual.Text = "Día Contable Actual :";
                lblMensaje.Text = "Para cerrar este Día Contable se debe cerrar el turno";
                IniciaPorDiaContable = true;
                DiaContableAperturado = true;
            }
        }
        else
        {
            // Modo Cierre
            lblDiaContableActual.Text = "Día Contable a Cerrar :";
            lblMensaje.Text = "Proceso de cierre del día contable";
            btnCerrar.Visible = true;
            btnCerrar.Enabled = true;
            IniciaPorDiaContable = true;
            DiaContableAperturado = true;
        }
    }

    private async Task OnAperturaClickAsync()
    {
        if (_aperturarHandler is null)
        {
            MessageBox.Show("Handler de apertura no configurado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _aperturarHandler.HandleAsync(
            new AperturarDiaContableCommand(dtpDiaContable.Value, _usuario));

        if (!result.EsExitoso)
        {
            // BR-DC-003: fecha menor al último día contable
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            dtpDiaContable.Focus();
            return;
        }

        // Legacy: lIniciaPorDiaContable = True; lDiaContableAperturado = True
        IniciaPorDiaContable = true;
        DiaContableAperturado = true;
        Close();
    }

    private async Task OnCerrarClickAsync()
    {
        if (MessageBox.Show(
                "¿Desea Cerrar el Día Contable?",
                Text,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        if (_cerrarHandler is null)
        {
            MessageBox.Show("Handler de cierre no configurado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _cerrarHandler.HandleAsync(
            new CerrarDiaContableCommand(dtpDiaContable.Value, _usuario));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        Close();
    }

    private void OnSalirClick(object? sender, EventArgs e)
    {
        // BR-DC-001: si no aperturó, el POS no puede vender
        if (!IniciaPorDiaContable)
            MessageBox.Show(
                "No puede Iniciar Venta sin Aperturar Día Contable",
                Text,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

        Close();
    }
}
