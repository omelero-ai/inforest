using Inforest.Application.Turno;
using Inforest.Desktop.Shared;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Desktop.Turno;

/// <summary>
/// Formulario de Apertura de Turno.
///
/// Legacy: <c>frmInicio.frm</c> — "Inicio de Turno".
///
/// Flujo:
///   1. Form_Load: consulta el último turno de la caja → determina si es apertura nueva o re-apertura.
///   2. Si hay moneda extranjera: carga TC del día y muestra controles de tipo de cambio.
///   3. Botones de Abono MN / ME: abren entrada numérica para ingresar montos iniciales.
///   4. "Aperturar": valida TC (si aplica) → confirma → llama AbrirTurnoHandler.
///      INSERT o UPDATE en MTURNO según corresponda (BR-TURNO-001, BR-TURNO-002).
///   5. "Cancelar": cierra sin acción.
///
/// Reglas: BR-TURNO-001, BR-TURNO-002, BR-TC-001.
/// Legacy: frmInicio.frm, cmdOpcion_Click(0), Form_Load.
/// SQL: MTURNO (INSERT/UPDATE), TTIPOCAMBIO (SELECT/spIns_TipoCambio).
/// </summary>
public class FrmAperturaTurno : Form
{
    // ── Colores del sistema ────────────────────────────────────────────────
    private static readonly Color ColorPrimario   = Color.FromArgb(0x3A, 0xAC, 0x6E);
    private static readonly Color ColorFondo      = Color.FromArgb(0xF8, 0xF9, 0xFC);
    private static readonly Color ColorValorRO    = Color.FromArgb(0xE0, 0xE0, 0xE0);
    private static readonly Color ColorValorMN    = Color.FromArgb(0x80, 0x80, 0x00);
    private static readonly Color ColorValorTotal = Color.FromArgb(0x80, 0x00, 0x00);

    // ── Dependencias ──────────────────────────────────────────────────────
    private readonly AbrirTurnoHandler _abrirTurnoHandler;
    private readonly ObtenerUltimoTurnoHandler _obtenerUltimoHandler;
    private readonly ITipoCambioRepository? _tipoCambioRepo;

    // ── Contexto de sesión ────────────────────────────────────────────────
    private readonly string _codigoCaja;
    private readonly string _codigoUsuario;
    private readonly string _codigoSalon;
    private readonly ModoConsultaTurno _modoTurno;

    // ── Configuración de monedas (de ConfiguracionSistema) ────────────────
    private readonly string _monedaN;   // sMonN — moneda nacional
    private readonly string _monedaE;   // sMonE
    private readonly string _pais;      // "000" Perú, "002" Colombia, "003" Bolivia

    // ── Estado interno ────────────────────────────────────────────────────
    private TurnoExistente? _turnoExistente;
    private bool _esReApertura;

    // Tipo de cambio
    private decimal _nTC;    // nVenta
    private decimal _nTC2;
    private decimal _nTC3;
    private decimal _nTCO;   // nOficial
    private bool _tcNuevo;   // wAgrega — true si hay que insertar TC

    // Montos
    private decimal _abonoN;
    private decimal _abonoE;
    private decimal _anteriorN;
    private decimal _anteriorE;

    // ── Controles UI ──────────────────────────────────────────────────────
    private Label lblFecha = null!;
    private GroupBox fraMontos = null!;

    // MN row
    private Label lblMN = null!;
    private Button btnAbonoN = null!;
    private TextBox txtAbonoN = null!;
    private Label txtAnteriorN = null!;
    private Label txtSaldoN = null!;

    // ME row (condicional)
    private Label? lblME;
    private Button? btnAbonoE;
    private TextBox? txtAbonoE;
    private Label? txtAnteriorE;
    private Label? txtSaldoE;

    // TC section
    private Panel panelTC = null!;
    private Button? btnTC;
    private Button? btnTCOficial;
    private Button? btnTC2;
    private Button? btnTC3;
    private Label? lblTC;
    private Label? lblTCO;
    private Label? lblTC2;
    private Label? lblTC3;

    // Main buttons
    private Button btnAperturar = null!;
    private Button btnCancelar = null!;

    /// <summary>Código del turno abierto, disponible tras DialogResult.OK.</summary>
    public string CodigoTurnoAbierto { get; private set; } = string.Empty;
    /// <summary>True si fue re-apertura de un turno existente.</summary>
    public bool FueReApertura { get; private set; }

    public FrmAperturaTurno(
        AbrirTurnoHandler abrirTurnoHandler,
        ObtenerUltimoTurnoHandler obtenerUltimoHandler,
        string codigoCaja,
        string codigoUsuario,
        string codigoSalon = "",
        ModoConsultaTurno modoTurno = ModoConsultaTurno.PorCaja,
        string monedaN = "S/.",
        string monedaE = "",
        string pais = "000",
        ITipoCambioRepository? tipoCambioRepo = null)
    {
        _abrirTurnoHandler    = abrirTurnoHandler;
        _obtenerUltimoHandler = obtenerUltimoHandler;
        _codigoCaja           = codigoCaja;
        _codigoUsuario        = codigoUsuario;
        _codigoSalon          = codigoSalon;
        _modoTurno            = modoTurno;
        _monedaN              = monedaN;
        _monedaE              = monedaE;
        _pais                 = pais;
        _tipoCambioRepo       = tipoCambioRepo;

        InitializeComponent();
        Load += async (_, _) => await OnLoadAsync();
    }

    // ──────────────────────────────────────────────────────────────────────
    // UI Setup
    // ──────────────────────────────────────────────────────────────────────
    private void InitializeComponent()
    {
        bool tieneMonedaE = TieneMonedaExtranjera();
        int alturaExtra   = tieneMonedaE ? 60 : 0; // fila ME
        int alturaTC      = tieneMonedaE ? 80 : 0; // sección TC

        Text             = "Apertura de Turno";
        StartPosition    = FormStartPosition.CenterScreen;
        FormBorderStyle  = FormBorderStyle.FixedSingle;
        MaximizeBox      = false;
        MinimizeBox      = false;
        BackColor        = ColorFondo;
        Font             = new Font("Segoe UI", 9f);
        ClientSize       = new Size(560, 320 + alturaExtra + alturaTC);

        // Fecha
        lblFecha = new Label
        {
            Left      = 20,  Top    = 14,
            Width     = 520, Height = 22,
            Text      = $"Fecha: {DateTime.Now:dd/MM/yyyy}",
            Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.Navy
        };

        // Frame Montos
        fraMontos = new GroupBox
        {
            Left      = 10,  Top    = 40,
            Width     = 540, Height = 140 + alturaExtra,
            Text      = " Montos ",
            Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            ForeColor = ColorPrimario,
            BackColor = ColorFondo
        };

        // Column headers dentro del frame
        var hAbono    = MakeLabel("Abono Actual",   210, 10, 110, bold: true);
        var hAnterior = MakeLabel("Saldo Anterior", 330, 10, 110, bold: true);
        var hInicial  = MakeLabel("Monto Inicial",  450, 10, 100, bold: true);
        hAbono.ForeColor = hAnterior.ForeColor = hInicial.ForeColor = Color.FromArgb(0x00, 0x73, 0xB4);
        fraMontos.Controls.AddRange([hAbono, hAnterior, hInicial]);

        // Row MN
        lblMN        = MakeLabel(_monedaN, 2, 38, 65);
        btnAbonoN    = MakeButton("Abono MN", 70,  30, 130, 32);
        txtAbonoN    = MakeMontoBox(210, 33, 110, readOnly: false);
        txtAnteriorN = MakeReadOnlyLabel(330, 36, 110);
        txtSaldoN    = MakeReadOnlyLabel(450, 36, 100, bold: true, color: ColorValorTotal);
        btnAbonoN.Click += (_, _) => IngresoNumerico("Abono MN", ref _abonoN, txtAbonoN, txtSaldoN, _anteriorN);
        fraMontos.Controls.AddRange([lblMN, btnAbonoN, txtAbonoN, txtAnteriorN, txtSaldoN]);

        // Row ME (condicional)
        if (tieneMonedaE)
        {
            lblME        = MakeLabel(_monedaE, 2, 95, 65);
            btnAbonoE    = MakeButton("Abono ME", 70, 87, 130, 32);
            txtAbonoE    = MakeMontoBox(210, 90, 110, readOnly: false);
            txtAnteriorE = MakeReadOnlyLabel(330, 93, 110);
            txtSaldoE    = MakeReadOnlyLabel(450, 93, 100, bold: true, color: ColorValorTotal);
            btnAbonoE.Click += (_, _) => IngresoNumerico("Abono ME", ref _abonoE, txtAbonoE!, txtSaldoE!, _anteriorE);
            fraMontos.Controls.AddRange([lblME, btnAbonoE, txtAbonoE, txtAnteriorE, txtSaldoE]);
        }

        // Panel TC
        panelTC = new Panel
        {
            Left      = 10,
            Top       = fraMontos.Bottom + 8,
            Width     = 540,
            Height    = alturaTC > 0 ? alturaTC : 4,
            Visible   = tieneMonedaE,
            BackColor = ColorFondo
        };

        if (tieneMonedaE)
        {
            btnTC = MakeButton("Tipo de Cambio", 0, 2, 140, 30);
            lblTC = MakeLabel("0.000", 148, 8, 100);
            lblTC.Font = new Font("Segoe UI", 9f, FontStyle.Bold);

            btnTCOficial = MakeButton("TC Oficial",  256, 2, 120, 30);
            lblTCO       = MakeLabel("0.000",        384, 8, 100);
            lblTCO.Font  = new Font("Segoe UI", 9f, FontStyle.Bold);

            btnTC.Click        += (_, _) => IngresoTC("TC Venta",   ref _nTC,  lblTC!);
            btnTCOficial.Click += (_, _) => IngresoTC("TC Oficial", ref _nTCO, lblTCO!);

            panelTC.Controls.AddRange([btnTC, lblTC, btnTCOficial, lblTCO]);

            if (_pais == "003")
            {
                btnTC2 = MakeButton("TC2", 0,   38, 120, 28);
                lblTC2 = MakeLabel("0.000", 128, 44,  90);
                btnTC3 = MakeButton("TC3", 224, 38, 120, 28);
                lblTC3 = MakeLabel("0.000", 352, 44,  90);
                btnTC2.Click += (_, _) => IngresoTC("Tipo Cambio 2", ref _nTC2, lblTC2!);
                btnTC3.Click += (_, _) => IngresoTC("Tipo Cambio 3", ref _nTC3, lblTC3!);
                panelTC.Controls.AddRange([btnTC2, lblTC2, btnTC3, lblTC3]);
            }
        }

        // Botones principales
        int btnTop = panelTC.Bottom + 16;
        btnAperturar = new Button
        {
            Left      = 200, Top    = btnTop,
            Width     = 120, Height = 38,
            Text      = "Aperturar",
            BackColor = ColorPrimario,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold)
        };
        btnCancelar = new Button
        {
            Left      = 340, Top    = btnTop,
            Width     = 120, Height = 38,
            Text      = "Cancelar",
            BackColor = Color.FromArgb(220, 220, 220),
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 9.5f)
        };

        btnAperturar.Click += async (_, _) => await AperturarAsync();
        btnCancelar.Click  += (_, _) => Close();

        Controls.AddRange([lblFecha, fraMontos, panelTC, btnAperturar, btnCancelar]);
        ClientSize = new Size(560, btnTop + 58);
    }

    // ──────────────────────────────────────────────────────────────────────
    // Load
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Legacy: frmInicio.frm Form_Load.
    /// Consulta el último turno de la caja y el TC del día.
    /// Determina si es apertura nueva o re-apertura.
    /// </summary>
    private async Task OnLoadAsync()
    {
        try
        {
            // Consultar último turno de la caja (incluye cerrados)
            var result = await _obtenerUltimoHandler.HandleAsync(
                new ObtenerUltimoTurnoQuery(_codigoCaja, _codigoUsuario, _modoTurno));

            _turnoExistente = result.EsExitoso ? result.Valor : null;

            if (_turnoExistente is null || _turnoExistente.Cerrado)
            {
                // Apertura nueva
                _esReApertura = false;
                Text          = "Apertura de Turno";
                lblFecha.Text = $"Fecha: {DateTime.Now:dd/MM/yyyy}";
                InicializarMontosEnCero();
            }
            else
            {
                // Re-apertura de turno no cerrado
                _esReApertura   = true;
                Text            = "Re Apertura de Turno";
                lblFecha.Text   = $"Fecha: {_turnoExistente.FechaInicial:dd/MM/yyyy}";
                _abonoN         = _turnoExistente.MontoInicialMN;
                _abonoE         = _turnoExistente.MontoInicialME;
                _anteriorN      = 0m;
                _anteriorE      = 0m;
                ActualizarMontosUI();
            }

            // Cargar tipo de cambio del día (si aplica)
            if (TieneMonedaExtranjera() && _tipoCambioRepo is not null)
            {
                var tc = await _tipoCambioRepo.ObtenerDelDiaAsync();
                if (tc is not null)
                {
                    _nTC     = tc.Venta;
                    _nTC2    = tc.Venta2;
                    _nTC3    = tc.Venta3;
                    _nTCO    = tc.Oficial;
                    _tcNuevo = false;
                    ActualizarTCUI();
                    OcultarBotonesTCSiYaTieneValor();
                }
                else
                {
                    _tcNuevo = true;
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al cargar datos de turno: {ex.Message}",
                "Apertura de Turno", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    // ──────────────────────────────────────────────────────────────────────
    // Aperturar
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Legacy: frmInicio.frm cmdOpcion_Click(0) — Case Is = 0.
    /// Valida TC, confirma, llama AbrirTurnoHandler.
    /// BR-TURNO-001, BR-TURNO-002, BR-TC-001.
    /// </summary>
    private async Task AperturarAsync()
    {
        // Validar tipo de cambio (si hay moneda extranjera)
        if (TieneMonedaExtranjera() && !ValidarTipoCambio())
            return;

        // Confirmación — Legacy: "Seguro de Aperturar el Turno?" / "Seguro de Re Aperturar el Turno?"
        string mensaje = _esReApertura ? "¿Seguro de Re Aperturar el Turno?" : "¿Seguro de Aperturar el Turno?";
        if (MessageBox.Show(mensaje, "Apertura de Turno",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
            return;

        btnAperturar.Enabled = false;

        try
        {
            var command = new AbrirTurnoCommand(
                CodigoCaja:          _codigoCaja,
                CodigoUsuario:       _codigoUsuario,
                MontoInicial:        _abonoN,
                MontoInicialME:      _abonoE,
                CodigoSalon:         _codigoSalon,
                ModoTurno:           _modoTurno,
                RegistrarTipoCambio: TieneMonedaExtranjera() && _tcNuevo,
                TipoCambio:          _nTC,
                TipoCambioOficial:   _nTCO,
                TipoCambio2:         _nTC2,
                TipoCambio3:         _nTC3);

            var result = await _abrirTurnoHandler.HandleAsync(command);

            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Apertura de Turno",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            CodigoTurnoAbierto = result.Valor!.CodigoTurno;
            FueReApertura      = result.Valor.EsReApertura;
            DialogResult       = DialogResult.OK;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al aperturar el turno: {ex.Message}",
                "Apertura de Turno", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnAperturar.Enabled = true;
        }
    }

    // ──────────────────────────────────────────────────────────────────────
    // Ingreso de valores (equivalente a frmNumPad)
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Legacy: frmNumPad.Show vbModal — ingreso de abono MN / ME.
    /// </summary>
    private void IngresoNumerico(string titulo, ref decimal valor, TextBox txtControl, Label lblSaldo, decimal anterior)
    {
        using var dlg = new FrmNumPad(titulo, valor);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            valor = dlg.Valor;
            txtControl.Text = valor.ToString("###,###,###,##0.00");
            lblSaldo.Text   = (valor + anterior).ToString("###,###,###,##0.00");
        }
    }

    /// <summary>
    /// Legacy: cmdOpcion(2/5/6) — ingreso de tipo de cambio con frmNumPad.
    /// </summary>
    private void IngresoTC(string titulo, ref decimal tcValor, Label lblMostrar)
    {
        using var dlg = new FrmNumPad(titulo, tcValor, decimales: 3);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            tcValor = dlg.Valor;
            lblMostrar.Text = tcValor.ToString("###,###,##0.000");
            _tcNuevo = true;
        }
    }

    // ──────────────────────────────────────────────────────────────────────
    // Helpers de validación y UI
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Legacy: frmInicio.frm — validar tipo de cambio antes de aperturar.
    /// pais="003": nTC, nTC2, nTC3 > 0.
    /// Otros: nTC > 0 si hay moneda extranjera.
    /// </summary>
    private bool ValidarTipoCambio()
    {
        if (_pais == "003")
        {
            if (_nTC == 0 || _nTC2 == 0 || _nTC3 == 0)
            {
                MessageBox.Show("Tipo de cambio no ingresado", "Apertura de Turno",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }
        else if (_nTC == 0)
        {
            MessageBox.Show("Tipo de cambio no ingresado", "Apertura de Turno",
                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            return false;
        }
        return true;
    }

    private bool TieneMonedaExtranjera()
        => !string.IsNullOrEmpty(_monedaE) && _monedaE != _monedaN;

    private void InicializarMontosEnCero()
    {
        _abonoN = 0m; _abonoE = 0m;
        _anteriorN = 0m; _anteriorE = 0m;
        ActualizarMontosUI();
    }

    private void ActualizarMontosUI()
    {
        txtAbonoN.Text    = _abonoN.ToString("###,###,##0.00");
        txtAnteriorN.Text = _anteriorN.ToString("###,###,##0.00");
        txtSaldoN.Text    = (_abonoN + _anteriorN).ToString("###,###,##0.00");
        if (TieneMonedaExtranjera() && txtAbonoE is not null)
        {
            txtAbonoE.Text    = _abonoE.ToString("###,###,##0.00");
            txtAnteriorE!.Text = _anteriorE.ToString("###,###,##0.00");
            txtSaldoE!.Text    = (_abonoE + _anteriorE).ToString("###,###,##0.00");
        }
    }

    private void ActualizarTCUI()
    {
        if (lblTC  is not null) lblTC.Text  = _nTC.ToString("###,###,##0.000");
        if (lblTCO is not null) lblTCO.Text = _nTCO.ToString("###,###,##0.000");
        if (lblTC2 is not null) lblTC2.Text = _nTC2.ToString("###,###,##0.000");
        if (lblTC3 is not null) lblTC3.Text = _nTC3.ToString("###,###,##0.000");
    }

    /// <summary>
    /// Legacy: ocultar botón Tipo de Cambio si ya está ingresado.
    /// </summary>
    private void OcultarBotonesTCSiYaTieneValor()
    {
        if (btnTC       is not null) btnTC.Visible       = _nTC  == 0;
        if (btnTCOficial is not null) btnTCOficial.Visible = _nTCO == 0;
        if (btnTC2      is not null) btnTC2.Visible      = _nTC2 == 0;
        if (btnTC3      is not null) btnTC3.Visible      = _nTC3 == 0;
    }

    // ──────────────────────────────────────────────────────────────────────
    // Factory helpers para controles
    // ──────────────────────────────────────────────────────────────────────

    private static Label MakeLabel(string text, int x, int y, int w, bool bold = false, Color? color = null)
        => new()
        {
            Text      = text,
            Left      = x, Top = y, Width = w, Height = 20,
            AutoSize  = false,
            Font      = bold ? new Font("Segoe UI", 8.5f, FontStyle.Bold) : new Font("Segoe UI", 8.5f),
            ForeColor = color ?? Color.FromArgb(0x23, 0x23, 0x23),
            BackColor = Color.Transparent
        };

    private static Button MakeButton(string text, int x, int y, int w, int h)
        => new()
        {
            Text      = text,
            Left      = x, Top = y, Width = w, Height = h,
            BackColor = Color.FromArgb(0xF2, 0xFA, 0xED),
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 8.5f)
        };

    private static TextBox MakeMontoBox(int x, int y, int w, bool readOnly)
        => new()
        {
            Left      = x, Top = y, Width = w, Height = 22,
            TextAlign = HorizontalAlignment.Right,
            BackColor = readOnly ? ColorValorRO : Color.White,
            ForeColor = ColorValorMN,
            Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            ReadOnly  = readOnly,
            Text      = "0.00"
        };

    private static Label MakeReadOnlyLabel(int x, int y, int w, bool bold = false, Color? color = null)
        => new()
        {
            Left        = x, Top = y, Width = w, Height = 20,
            BackColor   = ColorValorRO,
            BorderStyle = BorderStyle.FixedSingle,
            TextAlign   = ContentAlignment.MiddleRight,
            AutoSize    = false,
            Font        = new Font("Segoe UI", 8.5f, bold ? FontStyle.Bold : FontStyle.Regular),
            ForeColor   = color ?? ColorValorMN,
            Text        = "0.00"
        };
}

