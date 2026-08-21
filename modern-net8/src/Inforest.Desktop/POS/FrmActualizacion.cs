namespace Inforest.Desktop.POS;

/// <summary>
/// Formulario de notificación de nueva versión disponible.
/// Legacy: FrmActualizacion.frm — muestra tipo ("FE" o "INFOREST") y mensaje de actualización.
///
/// Reglas: BR-ACT-001.
/// SQL: TPARAMETROVERSION (VersionInfoRest, VersionFeInfoRest, VersionActualFEInfoRest).
/// </summary>
public class FrmActualizacion : Form
{
    // ── Colores de diseño (Guía de diseño estandarizado v1.0) ──────────────────
    private static readonly Color ColorPrimario = Color.FromArgb(0x3A, 0xAC, 0x6E);
    private static readonly Color ColorFondo    = Color.FromArgb(0xF8, 0xF9, 0xFC);
    private static readonly Color ColorTexto    = Color.FromArgb(0x23, 0x23, 0x23);

    private readonly Label _lblTipo;
    private readonly Label _lblMensaje;
    private readonly Button _btnCerrar;

    /// <summary>Tipo de actualización: "INFOREST" o "FE".</summary>
    public string Tipo { get; set; } = "INFOREST";

    public FrmActualizacion(string mensaje)
    {
        // ── Controles ──────────────────────────────────────────────────────────
        _lblTipo = new Label
        {
            AutoSize  = false,
            Dock      = DockStyle.Top,
            Height    = 36,
            TextAlign = ContentAlignment.MiddleCenter,
            Font      = new Font("Segoe UI", 13f, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = ColorPrimario,
            Padding   = new Padding(8, 4, 8, 4)
        };

        _lblMensaje = new Label
        {
            AutoSize  = false,
            Dock      = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Font      = new Font("Segoe UI", 10f),
            ForeColor = ColorTexto,
            Text      = mensaje,
            Padding   = new Padding(16)
        };

        _btnCerrar = new Button
        {
            Text      = "Cerrar",
            Dock      = DockStyle.Bottom,
            Height    = 38,
            BackColor = ColorPrimario,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor    = Cursors.Hand
        };
        _btnCerrar.FlatAppearance.BorderSize = 0;
        _btnCerrar.Click += (_, _) => Close();

        // ── Formulario ─────────────────────────────────────────────────────────
        Text            = "Actualización Disponible";
        Size            = new Size(480, 240);
        StartPosition   = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox     = false;
        MaximizeBox     = false;
        BackColor       = ColorFondo;

        Controls.Add(_lblMensaje);
        Controls.Add(_lblTipo);
        Controls.Add(_btnCerrar);

        Load += (_, _) => _lblTipo.Text = Tipo == "FE"
            ? "Actualización de Facturación Electrónica"
            : "Actualización de INFOREST";
    }
}
