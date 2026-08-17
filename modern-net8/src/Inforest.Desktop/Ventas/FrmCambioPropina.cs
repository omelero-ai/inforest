using Inforest.Desktop.Shared;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Diálogo modal de cambio de propina (tip).
/// Legacy: frmCambioPropina.frm — mostrado desde frmPedidoCorrelativo y frmDocumentoCorrelativo.
/// Permite seleccionar propina en moneda nacional (MN) o moneda extranjera (ME).
/// BR-PROPINA-001: propina MN tMonedaPropina='01'; BR-PROPINA-002: propina ME tMonedaPropina='02'.
/// BR-PROPINA-003: solo una moneda activa a la vez; BR-PROPINA-004: valor 0 = sin propina.
/// </summary>
public sealed class FrmCambioPropina : Form
{
    // ── Resultado ────────────────────────────────────────────────────────────
    /// <summary>Monto de propina en moneda nacional. 0 si se eligió ME.</summary>
    public decimal PropinaMN   { get; private set; }

    /// <summary>Monto de propina en moneda extranjera. 0 si se eligió MN.</summary>
    public decimal PropinaME   { get; private set; }

    /// <summary>
    /// Tipo de propina seleccionado.
    /// "01" = moneda nacional, "02" = moneda extranjera, "" = sin propina.
    /// Equivalente a <c>tPropina</c> en VB6.
    /// </summary>
    public string  TipoPropina { get; private set; } = "";

    /// <summary>
    /// True si el usuario confirmó la selección (wFlag = True en VB6).
    /// </summary>
    public bool Confirmado { get; private set; }

    // ── Controles ────────────────────────────────────────────────────────────
    private readonly Button _btnPropinaMN;
    private readonly Button _btnPropinaME;
    private readonly Button _btnGrabar;
    private readonly string _simboloMN;
    private readonly string _simboloME;

    /// <summary>
    /// Crea el diálogo de cambio de propina.
    /// </summary>
    /// <param name="simboloMN">Símbolo de moneda nacional (ej. "S/.").</param>
    /// <param name="simboloME">Símbolo de moneda extranjera (ej. "US$").</param>
    /// <param name="propinaMNActual">Monto propina MN ya registrada (pre-carga el botón).</param>
    /// <param name="propinaMEActual">Monto propina ME ya registrada (pre-carga el botón).</param>
    /// <param name="tipoActual">Tipo propina actual: "01", "02" o vacío.</param>
    public FrmCambioPropina(
        string  simboloMN       = "S/.",
        string  simboloME       = "US$",
        decimal propinaMNActual = 0,
        decimal propinaMEActual = 0,
        string  tipoActual      = "")
    {
        _simboloMN  = simboloMN;
        _simboloME  = simboloME;
        PropinaMN   = propinaMNActual;
        PropinaME   = propinaMEActual;
        TipoPropina = tipoActual;

        Text            = "Cambio de Propina";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition   = FormStartPosition.CenterParent;
        MaximizeBox     = false;
        MinimizeBox     = false;
        Width           = 520;
        Height          = 150;
        BackColor       = Color.FromArgb(245, 246, 247);

        // ── Título ─────────────────────────────────────────────────────────
        var titulo = new Label
        {
            Text      = "Seleccione moneda de propina",
            Dock      = DockStyle.Top,
            Height    = 32,
            Font      = new Font("Segoe UI", 9, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding   = new Padding(10, 0, 0, 0),
            BackColor = Color.FromArgb(40, 60, 90),
            ForeColor = Color.White
        };

        // ── Botones propina ────────────────────────────────────────────────
        _btnPropinaMN = new Button
        {
            Text      = TextoPropina(simboloMN, propinaMNActual),
            Height    = 52,
            Dock      = DockStyle.None,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor    = Cursors.Hand,
            Margin    = new Padding(6)
        };
        _btnPropinaMN.FlatAppearance.BorderSize = 1;
        _btnPropinaMN.Click += (_, _) => IngresarPropina(esMN: true);

        _btnPropinaME = new Button
        {
            Text      = TextoPropina(simboloME, propinaMEActual),
            Height    = 52,
            Dock      = DockStyle.None,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor    = Cursors.Hand,
            Margin    = new Padding(6)
        };
        _btnPropinaME.FlatAppearance.BorderSize = 1;
        _btnPropinaME.Click += (_, _) => IngresarPropina(esMN: false);

        // ── Botones acción ─────────────────────────────────────────────────
        _btnGrabar = new Button
        {
            Text      = "Grabar",
            Height    = 52,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(40, 130, 70),
            ForeColor = Color.White,
            Font      = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor    = Cursors.Hand,
            Enabled   = tipoActual != "",   // BR-PROPINA-004: requiere al menos una selección
            Margin    = new Padding(6)
        };
        _btnGrabar.FlatAppearance.BorderSize = 0;
        _btnGrabar.Click += (_, _) =>
        {
            Confirmado   = true;
            DialogResult = DialogResult.OK;
            Close();
        };

        var btnSalir = new Button
        {
            Text      = "Salir",
            Height    = 52,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            Font      = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor    = Cursors.Hand,
            Margin    = new Padding(6)
        };
        btnSalir.FlatAppearance.BorderSize = 0;
        btnSalir.Click += (_, _) =>
        {
            Confirmado   = false;
            DialogResult = DialogResult.Cancel;
            Close();
        };

        // ── Layout ─────────────────────────────────────────────────────────
        var panel = new TableLayoutPanel
        {
            Dock        = DockStyle.Fill,
            ColumnCount = 4,
            RowCount    = 1,
            Padding     = new Padding(6, 4, 6, 4)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30f));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30f));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        panel.Controls.Add(_btnPropinaMN); panel.SetCellPosition(_btnPropinaMN, new TableLayoutPanelCellPosition(0, 0));
        panel.Controls.Add(_btnPropinaME); panel.SetCellPosition(_btnPropinaME, new TableLayoutPanelCellPosition(1, 0));
        panel.Controls.Add(_btnGrabar);    panel.SetCellPosition(_btnGrabar,    new TableLayoutPanelCellPosition(2, 0));
        panel.Controls.Add(btnSalir);      panel.SetCellPosition(btnSalir,      new TableLayoutPanelCellPosition(3, 0));

        Controls.Add(panel);
        Controls.Add(titulo);

        ActualizarEstilosBotones();
    }

    // ── Lógica ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Abre el NumPad para ingresar el monto de propina en la moneda indicada.
    /// BR-PROPINA-001/002.
    /// </summary>
    private void IngresarPropina(bool esMN)
    {
        var valorActual = esMN ? PropinaMN : PropinaME;
        using var dlg = new FrmNumPad(
            titulo        : $"Propina {(esMN ? _simboloMN : _simboloME)}",
            valorInicial  : valorActual);

        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        if (esMN)
        {
            // BR-PROPINA-003: una sola moneda activa
            PropinaMN   = dlg.Valor;
            PropinaME   = 0;
            TipoPropina = "01";
            _btnPropinaMN.Text = TextoPropina(_simboloMN, PropinaMN);
            _btnPropinaME.Text = TextoPropina(_simboloME, 0);
        }
        else
        {
            PropinaME   = dlg.Valor;
            PropinaMN   = 0;
            TipoPropina = "02";
            _btnPropinaME.Text = TextoPropina(_simboloME, PropinaME);
            _btnPropinaMN.Text = TextoPropina(_simboloMN, 0);
        }

        // BR-PROPINA-004: habilitar Grabar si hay un monto > 0
        _btnGrabar.Enabled = (PropinaMN > 0 || PropinaME > 0);
        ActualizarEstilosBotones();
    }

    /// <summary>Refleja visualmente qué moneda tiene el monto activo (negrita vs normal).</summary>
    private void ActualizarEstilosBotones()
    {
        bool mnActiva = TipoPropina == "01" && PropinaMN > 0;
        bool meActiva = TipoPropina == "02" && PropinaME > 0;

        _btnPropinaMN.BackColor = mnActiva
            ? Color.FromArgb(40, 100, 170)
            : Color.FromArgb(70, 90, 130);
        _btnPropinaMN.ForeColor = Color.White;

        _btnPropinaME.BackColor = meActiva
            ? Color.FromArgb(40, 100, 170)
            : Color.FromArgb(70, 90, 130);
        _btnPropinaME.ForeColor = Color.White;
    }

    private static string TextoPropina(string simbolo, decimal monto)
        => $"Propina {simbolo} {monto:N2}";
}
