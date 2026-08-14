namespace Inforest.Desktop.POS;

/// <summary>
/// Teclado virtual en pantalla para captura de contraseña de supervisor.
/// Legacy: frmPassword.frm. BR-006.
/// Se usa como diálogo modal; el resultado queda en <see cref="Contrasena"/> si el usuario confirmó.
/// </summary>
public class FrmPassword : Form
{
    private readonly TextBox _txtMostrado;
    private string _valorInterno = string.Empty;

    /// <summary>Contraseña ingresada; solo disponible si <see cref="DialogResult"/> == OK.</summary>
    public string Contrasena => _valorInterno;

    public FrmPassword()
    {
        Text = "Ingrese clave";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(440, 360);

        // ── Display ────────────────────────────────────────────────────────────
        _txtMostrado = new TextBox
        {
            Left = 10, Top = 10, Width = 420,
            ReadOnly = true,
            UseSystemPasswordChar = true,
            Font = new Font("Courier New", 14)
        };
        Controls.Add(_txtMostrado);

        // ── Key layout (same as VB6 frmPassword) ──────────────────────────────
        string[] rows =
        [
            "! ' # $ % & / ( ) =",
            "Q W E R T Y U I O P",
            "A S D F G H J K L Ñ",
            "Z X C V B N M , . -",
            "0 1 2 3 4 5 6 7 8 9"
        ];

        int topStart = 50;
        int btnW = 38, btnH = 38, gap = 2;

        foreach (var row in rows)
        {
            var keys = row.Split(' ');
            int left = 10;
            foreach (var key in keys)
            {
                var btn = new Button { Text = key, Width = btnW, Height = btnH, Left = left, Top = topStart };
                var capturedKey = key;
                btn.Click += (_, _) =>
                {
                    _valorInterno += capturedKey == "&&" ? "&" : capturedKey;
                    _txtMostrado.Text = new string('*', _valorInterno.Length);
                };
                Controls.Add(btn);
                left += btnW + gap;
            }
            topStart += btnH + gap;
        }

        // ── Special buttons ────────────────────────────────────────────────────
        var btnSpace = new Button { Text = "SPACE", Width = 80, Height = btnH, Left = 10, Top = topStart };
        btnSpace.Click += (_, _) =>
        {
            _valorInterno += " ";
            _txtMostrado.Text = new string('*', _valorInterno.Length);
        };

        var btnBack = new Button { Text = "⌫ Back", Width = 80, Height = btnH, Left = 100, Top = topStart };
        btnBack.Click += (_, _) =>
        {
            if (_valorInterno.Length > 0)
            {
                _valorInterno = _valorInterno[..^1];
                _txtMostrado.Text = new string('*', _valorInterno.Length);
            }
        };

        var btnEnter = new Button { Text = "Enter ✓", Width = 100, Height = btnH, Left = 230, Top = topStart };
        btnEnter.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };

        var btnEsc = new Button { Text = "ESC ✗", Width = 80, Height = btnH, Left = 350, Top = topStart };
        btnEsc.Click += (_, _) => { _valorInterno = string.Empty; DialogResult = DialogResult.Cancel; Close(); };

        Controls.AddRange([btnSpace, btnBack, btnEnter, btnEsc]);
        ClientSize = new Size(440, topStart + btnH + 10);
    }
}
