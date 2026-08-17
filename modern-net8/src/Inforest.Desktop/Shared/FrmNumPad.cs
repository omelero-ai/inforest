namespace Inforest.Desktop.Shared;

/// <summary>
/// Teclado numérico modal reutilizable.
/// Legacy: frmNumPad.frm — usado en todo el sistema para capturar valores numéricos.
/// Equivalente de: sDescrip (VB6) = <see cref="ValorTexto"/>; wEnter (VB6) = DialogResult.OK.
/// BR-NUMPAD-001.
/// </summary>
public sealed class FrmNumPad : Form
{
    // ── Estado ───────────────────────────────────────────────────────────────
    private string _entrada = "0";
    private bool   _puntoPuesto;

    // ── Resultado ────────────────────────────────────────────────────────────
    /// <summary>Valor numérico ingresado (0 si el usuario canceló).</summary>
    public decimal Valor     { get; private set; }

    /// <summary>Texto exacto ingresado (equivalente a <c>sDescrip</c> en VB6).</summary>
    public string  ValorTexto { get; private set; } = "0";

    // ── Controles ────────────────────────────────────────────────────────────
    private readonly Label  _lblDisplay;
    private readonly Label  _lblTitulo;

    /// <summary>
    /// Crea un teclado numérico modal.
    /// </summary>
    /// <param name="titulo">Título del diálogo.</param>
    /// <param name="valorInicial">Valor pre-cargado (0 = vacío).</param>
    public FrmNumPad(string titulo = "Ingrese el valor", decimal valorInicial = 0)
    {
        Text            = titulo;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition   = FormStartPosition.CenterParent;
        MaximizeBox     = false;
        MinimizeBox     = false;
        Width           = 300;
        Height          = 380;
        BackColor       = Color.FromArgb(245, 246, 247);
        KeyPreview      = true;

        if (valorInicial != 0)
        {
            _entrada    = valorInicial.ToString("0.##");
            _puntoPuesto = _entrada.Contains('.');
        }

        // ── Título ─────────────────────────────────────────────────────────
        _lblTitulo = new Label
        {
            Text      = titulo,
            Dock      = DockStyle.Top,
            Height    = 36,
            Font      = new Font("Segoe UI", 9, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding   = new Padding(10, 0, 0, 0),
            BackColor = Color.FromArgb(40, 60, 90),
            ForeColor = Color.White
        };

        // ── Display ────────────────────────────────────────────────────────
        _lblDisplay = new Label
        {
            Text      = FormatearDisplay(_entrada),
            Height    = 50,
            Dock      = DockStyle.Top,
            Font      = new Font("Segoe UI", 16, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleRight,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Padding   = new Padding(0, 0, 8, 0)
        };

        // ── Teclado ────────────────────────────────────────────────────────
        var table = new TableLayoutPanel
        {
            Dock        = DockStyle.Fill,
            ColumnCount = 3,
            RowCount    = 5,
            Padding     = new Padding(6)
        };
        for (int i = 0; i < 3; i++) table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
        for (int i = 0; i < 5; i++) table.RowStyles.Add(new RowStyle(SizeType.Percent, 20f));

        // Fila 0: 7 8 9
        AgregarBotonNum(table, "7", 0, 0);
        AgregarBotonNum(table, "8", 1, 0);
        AgregarBotonNum(table, "9", 2, 0);
        // Fila 1: 4 5 6
        AgregarBotonNum(table, "4", 0, 1);
        AgregarBotonNum(table, "5", 1, 1);
        AgregarBotonNum(table, "6", 2, 1);
        // Fila 2: 1 2 3
        AgregarBotonNum(table, "1", 0, 2);
        AgregarBotonNum(table, "2", 1, 2);
        AgregarBotonNum(table, "3", 2, 2);
        // Fila 3: . 0 ⌫
        AgregarBotonEspecial(table, ".",  0, 3, Color.FromArgb(100, 120, 160));
        AgregarBotonNum(table, "0",  1, 3);
        AgregarBotonEspecial(table, "⌫", 2, 3, Color.FromArgb(180, 80, 60));
        // Fila 4: Limpiar | Enter
        var btnLimpiar = CrearBoton("Limpiar", Color.FromArgb(150, 150, 170));
        btnLimpiar.Click += (_, _) => { _entrada = "0"; _puntoPuesto = false; ActualizarDisplay(); };
        table.Controls.Add(btnLimpiar); table.SetCellPosition(btnLimpiar, new TableLayoutPanelCellPosition(0, 4));

        var btnEnter = CrearBoton("✔ Aceptar", Color.FromArgb(40, 130, 70));
        btnEnter.Click += (_, _) => Confirmar();
        table.Controls.Add(btnEnter); table.SetCellPosition(btnEnter, new TableLayoutPanelCellPosition(1, 4));
        table.SetColumnSpan(btnEnter, 2);

        Controls.Add(table);
        Controls.Add(_lblDisplay);
        Controls.Add(_lblTitulo);

        KeyDown += FrmNumPad_KeyDown;
        ActualizarDisplay();
    }

    // ── Teclado físico ───────────────────────────────────────────────────────
    private void FrmNumPad_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9)
            AgregarDigito(((char)('0' + (e.KeyCode - Keys.D0))).ToString());
        else if (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9)
            AgregarDigito(((char)('0' + (e.KeyCode - Keys.NumPad0))).ToString());
        else if (e.KeyCode == Keys.Decimal || e.KeyCode == Keys.OemPeriod)
            AgregarPunto();
        else if (e.KeyCode == Keys.Back)
            Borrar();
        else if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
            Confirmar();
        else if (e.KeyCode == Keys.Escape)
        { DialogResult = DialogResult.Cancel; Close(); }
        e.Handled = true;
    }

    // ── Helpers ──────────────────────────────────────────────────────────────
    private void AgregarBotonNum(TableLayoutPanel t, string digito, int col, int row)
    {
        var btn = CrearBoton(digito, Color.FromArgb(60, 80, 110));
        btn.Font = new Font("Segoe UI", 14, FontStyle.Bold);
        btn.Click += (_, _) => AgregarDigito(digito);
        t.Controls.Add(btn); t.SetCellPosition(btn, new TableLayoutPanelCellPosition(col, row));
    }

    private void AgregarBotonEspecial(TableLayoutPanel t, string label, int col, int row, Color color)
    {
        var btn = CrearBoton(label, color);
        btn.Font = new Font("Segoe UI", 12, FontStyle.Bold);
        btn.Click += (_, _) =>
        {
            if (label == ".") AgregarPunto();
            else Borrar();
        };
        t.Controls.Add(btn); t.SetCellPosition(btn, new TableLayoutPanelCellPosition(col, row));
    }

    private static Button CrearBoton(string text, Color backColor) => new()
    {
        Text       = text,
        Dock       = DockStyle.Fill,
        FlatStyle  = FlatStyle.Flat,
        BackColor  = backColor,
        ForeColor  = Color.White,
        Font       = new Font("Segoe UI", 10, FontStyle.Bold),
        Cursor     = Cursors.Hand,
        Margin     = new Padding(3)
    };

    private void AgregarDigito(string d)
    {
        if (_entrada == "0") _entrada = d;
        else _entrada += d;
        ActualizarDisplay();
    }

    private void AgregarPunto()
    {
        if (_puntoPuesto) return;
        _puntoPuesto = true;
        if (_entrada == "0" || _entrada == "") _entrada = "0.";
        else _entrada += ".";
        ActualizarDisplay();
    }

    private void Borrar()
    {
        if (_entrada.Length <= 1) { _entrada = "0"; _puntoPuesto = false; }
        else
        {
            if (_entrada[^1] == '.') _puntoPuesto = false;
            _entrada = _entrada[..^1];
        }
        ActualizarDisplay();
    }

    private void Confirmar()
    {
        ValorTexto  = _entrada;
        Valor       = decimal.TryParse(_entrada, out var v) ? v : 0;
        DialogResult = DialogResult.OK;
        Close();
    }

    private void ActualizarDisplay() => _lblDisplay.Text = FormatearDisplay(_entrada);

    private static string FormatearDisplay(string entrada)
    {
        if (decimal.TryParse(entrada, out var v) && !entrada.EndsWith('.'))
            return v.ToString("N2");
        return entrada;
    }
}
