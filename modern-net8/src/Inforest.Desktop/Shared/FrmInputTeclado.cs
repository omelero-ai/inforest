namespace Inforest.Desktop.Shared;

/// <summary>
/// Teclado alfanumérico en pantalla para entrada táctil de texto.
/// Legacy: frmKeyBoard.frm (usuario) y frmPassword.frm (contraseña).
/// Equivalente de: sDescrip (VB6) = <see cref="ValorTexto"/>; wEnter (VB6) = DialogResult.OK.
/// Usado en <c>frmAcceso.frm</c> imgOpcion Index=2 (usuario) e Index=3 (password).
/// </summary>
public sealed class FrmInputTeclado : Form
{
    // ── Resultado ────────────────────────────────────────────────────────────
    /// <summary>Texto ingresado. Vacío si el usuario canceló.</summary>
    public string ValorTexto { get; private set; } = string.Empty;

    // ── Controles / estado ───────────────────────────────────────────────────
    private readonly TextBox _txtDisplay;
    private readonly bool    _esContrasena;

    // Layout de teclas alfanuméricas — legacy frmKeyBoard incluía letras mayúsculas + dígitos
    private static readonly string[] FilasTeclas =
    [
        "1 2 3 4 5 6 7 8 9 0",
        "Q W E R T Y U I O P",
        "A S D F G H J K L Ñ",
        "Z X C V B N M . - _",
        "ESPACIO ⌫ OK CANCELAR"
    ];

    /// <param name="titulo">Título del diálogo.</param>
    /// <param name="textoInicial">Texto pre-cargado en el campo.</param>
    /// <param name="esContrasena">Cuando es true enmascara el texto con asteriscos.</param>
    public FrmInputTeclado(
        string titulo       = "Ingrese el texto",
        string textoInicial = "",
        bool   esContrasena = false)
    {
        _esContrasena     = esContrasena;
        Text              = titulo;
        FormBorderStyle   = FormBorderStyle.FixedDialog;
        StartPosition     = FormStartPosition.CenterParent;
        MaximizeBox       = false;
        MinimizeBox       = false;
        Width             = 560;
        Height            = 390;
        BackColor         = Color.FromArgb(245, 246, 247);
        KeyPreview        = true;
        KeyDown          += FrmInputTeclado_KeyDown;

        // ── Header ─────────────────────────────────────────────────────────
        var lblTitulo = new Label
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
        _txtDisplay = new TextBox
        {
            Left                  = 10,
            Top                   = 46,
            Width                 = 530,
            Height                = 36,
            Font                  = new Font("Segoe UI", 14, FontStyle.Bold),
            BackColor             = Color.White,
            BorderStyle           = BorderStyle.FixedSingle,
            UseSystemPasswordChar = esContrasena,
            CharacterCasing       = esContrasena ? CharacterCasing.Normal : CharacterCasing.Upper,
            Text                  = esContrasena ? string.Empty : textoInicial.ToUpperInvariant(),
            ReadOnly              = true
        };

        Controls.Add(lblTitulo);
        Controls.Add(_txtDisplay);

        // ── Teclado ────────────────────────────────────────────────────────
        int y = 92;
        foreach (var fila in FilasTeclas)
        {
            var teclas = fila.Split(' ');
            int x = 10;

            foreach (var tecla in teclas)
            {
                int w = tecla switch
                {
                    "ESPACIO"  => 160,
                    "OK"       => 80,
                    "CANCELAR" => 80,
                    "⌫"        => 80,
                    _          => 48
                };

                var btn = new Button
                {
                    Text      = tecla,
                    Left      = x,
                    Top       = y,
                    Width     = w,
                    Height    = 40,
                    Font      = new Font("Segoe UI", 10),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = tecla switch
                    {
                        "OK"       => Color.FromArgb(66, 198, 126),
                        "CANCELAR" => Color.FromArgb(220, 80, 80),
                        _          => Color.White
                    },
                    ForeColor = tecla is "OK" or "CANCELAR" ? Color.White : Color.Black,
                    Tag       = tecla
                };
                btn.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
                btn.Click += Tecla_Click;
                Controls.Add(btn);
                x += w + 4;
            }

            y += 44;
        }
    }

    private void Tecla_Click(object? sender, EventArgs e)
    {
        if (sender is not Button btn) return;
        var tecla = btn.Tag as string ?? string.Empty;

        switch (tecla)
        {
            case "OK":
                ValorTexto   = _txtDisplay.Text;
                DialogResult = DialogResult.OK;
                Close();
                break;

            case "CANCELAR":
                ValorTexto   = string.Empty;
                DialogResult = DialogResult.Cancel;
                Close();
                break;

            case "⌫":
                if (_txtDisplay.Text.Length > 0)
                    _txtDisplay.Text = _txtDisplay.Text[..^1];
                break;

            case "ESPACIO":
                _txtDisplay.Text += ' ';
                break;

            default:
                _txtDisplay.Text += _esContrasena ? tecla : tecla.ToUpperInvariant();
                break;
        }
    }

    private void FrmInputTeclado_KeyDown(object? sender, KeyEventArgs e)
    {
        switch (e.KeyCode)
        {
            case Keys.Enter:
                ValorTexto   = _txtDisplay.Text;
                DialogResult = DialogResult.OK;
                Close();
                break;

            case Keys.Escape:
                ValorTexto   = string.Empty;
                DialogResult = DialogResult.Cancel;
                Close();
                break;

            case Keys.Back:
                if (_txtDisplay.Text.Length > 0)
                    _txtDisplay.Text = _txtDisplay.Text[..^1];
                break;
        }
    }
}
