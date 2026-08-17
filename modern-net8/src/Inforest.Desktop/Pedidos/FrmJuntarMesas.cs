using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Pedidos;

/// <summary>
/// Resultado devuelto por <see cref="FrmJuntarMesas"/> al caller.
/// Legacy: frmJuntarMesas.frm — picGuardar_Click devuelve mesaOriginal + filtroMesasAdicionales al frmMesas.
/// </summary>
public sealed record JuntarMesasResult(
    bool Confirmado,
    string MesaOriginal,
    IReadOnlyList<string> MesasAdicionales);

/// <summary>
/// Selector visual de mesas para unificar varios mesas bajo un mismo pedido.
/// Legacy: frmJuntarMesas.frm
/// Reglas: BR-JUNTAR-001, BR-JUNTAR-002, BR-JUNTAR-003.
///
/// BR-JUNTAR-001: Solo se pueden seleccionar mesas en estado Libre; la mesa original
///                siempre queda preseleccionada y no puede deseleccionarse.
/// BR-JUNTAR-002: El formulario no persiste datos — devuelve el resultado al caller
///                a través de <see cref="Resultado"/>.
/// BR-JUNTAR-003: Al confirmar se retorna la mesa original más la lista de mesas adicionales.
/// </summary>
public sealed class FrmJuntarMesas : Form
{
    private readonly IMesaRepository _mesaRepository;
    private readonly string _codigoSalon;
    private readonly string _codigoMesaOriginal;
    private readonly int _numAdultos;
    private readonly int _numNinos;

    // Estado interno de selección
    private readonly HashSet<string> _mesasSeleccionadas = [];

    // Controles dinámicos para las mesas
    private readonly List<(Button Btn, Mesa Mesa)> _botonesMesa = [];

    // Controles fijos
    private Label _lblSalon = null!;
    private Label _lblMesasSeleccionadas = null!;
    private Label _lblConteo = null!;
    private FlowLayoutPanel _flpMesas = null!;

    // Colores (migrados desde las constantes vbHex* del Legacy)
    private static readonly Color ColorLibre           = Color.FromArgb(198, 229, 206);   // vbHexLista    "#01" verde claro
    private static readonly Color ColorOcupada         = Color.FromArgb(229, 155, 116);   // vbHexOcupada  "#02" naranja
    private static readonly Color ColorReservada       = Color.FromArgb(176, 196, 222);   // vbHexReservada"#03" azul claro
    private static readonly Color ColorBloqueada       = Color.FromArgb(211, 211, 211);   // vbHexBloqueada"#05" gris
    private static readonly Color ColorOriginal        = Color.FromArgb(118, 228, 230);   // vbHexMesasjuntas "#76E4E6"
    private static readonly Color ColorSeleccionBorde  = Color.FromArgb(64,  22,  13);    // tcolorSeleccion  "#40160d"
    private static readonly Color ColorNormalBorde     = Color.FromArgb(207, 211, 216);   // tcolorDeseleccion "#CFD3D8"

    /// <summary>Resultado de la operación; se establece antes de cerrar el formulario.</summary>
    public JuntarMesasResult Resultado { get; private set; } =
        new JuntarMesasResult(false, string.Empty, []);

    public FrmJuntarMesas(
        IMesaRepository mesaRepository,
        string codigoSalon,
        string codigoMesaOriginal,
        int numAdultos = 0,
        int numNinos = 0)
    {
        _mesaRepository = mesaRepository;
        _codigoSalon = codigoSalon;
        _codigoMesaOriginal = codigoMesaOriginal;
        _numAdultos = numAdultos;
        _numNinos = numNinos;

        Text = "Juntar Mesas";
        Size = new Size(860, 580);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        BuildLayout();
        Load += async (_, _) => await CargarMesasAsync();
    }

    // ── Layout ──────────────────────────────────────────────────────────────────

    private void BuildLayout()
    {
        // Header azul
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 36,
            BackColor = Color.FromArgb(0, 84, 166)
        };
        _lblSalon = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Fill,
            Padding = new Padding(10, 8, 0, 0),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.White,
            Text = "Juntar Mesas"
        };
        pnlHeader.Controls.Add(_lblSalon);

        // Footer con info y botones
        var pnlFooter = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 86
        };
        BuildFooter(pnlFooter);

        // Panel central con mesas
        var pnlCentro = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8) };
        _flpMesas = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Padding = new Padding(4)
        };
        pnlCentro.Controls.Add(_flpMesas);

        Controls.Add(pnlCentro);
        Controls.Add(pnlFooter);
        Controls.Add(pnlHeader);
    }

    private void BuildFooter(Panel pnlFooter)
    {
        var lblMesasLabel = new Label
        {
            AutoSize = false,
            Left = 10, Top = 10,
            Width = 140, Height = 20,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            Text = "Número de mesas juntas:"
        };
        _lblConteo = new Label
        {
            AutoSize = false,
            Left = 155, Top = 10,
            Width = 50, Height = 20,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 84, 166),
            Text = "0"
        };
        var lblMesasNombresLabel = new Label
        {
            AutoSize = false,
            Left = 10, Top = 34,
            Width = 100, Height = 20,
            Font = new Font("Segoe UI", 8.5f),
            Text = "Mesas a juntar:"
        };
        _lblMesasSeleccionadas = new Label
        {
            AutoSize = false,
            Left = 115, Top = 34,
            Width = 500, Height = 20,
            Font = new Font("Segoe UI", 8.5f),
            ForeColor = Color.DarkSlateGray,
            Text = ""
        };

        var btnGuardar = new Button
        {
            Text = "Confirmar",
            Width = 100, Height = 30,
            Left = pnlFooter.Width - 230,
            Top = 28,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
            BackColor = Color.FromArgb(0, 84, 166),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        btnGuardar.Click += OnConfirmar;

        var btnCancelar = new Button
        {
            Text = "Cancelar",
            Width = 100, Height = 30,
            Left = pnlFooter.Width - 120,
            Top = 28,
            Anchor = AnchorStyles.Right | AnchorStyles.Top
        };
        btnCancelar.Click += (_, _) => Close();

        pnlFooter.Controls.AddRange([lblMesasLabel, _lblConteo, lblMesasNombresLabel, _lblMesasSeleccionadas, btnGuardar, btnCancelar]);
    }

    // ── Carga de mesas ──────────────────────────────────────────────────────────

    private async Task CargarMesasAsync()
    {
        var todasMesas = await _mesaRepository.ObtenerTodosAsync();
        var mesasSalon = todasMesas
            .Where(m => m.Activo && m.CodigoSalon == _codigoSalon)
            .OrderBy(m => m.CodigoMesa)
            .ToList();

        // Pre-seleccionar la mesa original
        _mesasSeleccionadas.Add(_codigoMesaOriginal);

        // Buscar nombre del salón
        var mesaOrig = mesasSalon.FirstOrDefault(m => m.CodigoMesa == _codigoMesaOriginal);
        _lblSalon.Text = mesaOrig is not null
            ? $"Juntar Mesas — Salón: {_codigoSalon}  |  Mesa original: {mesaOrig.Resumido}"
            : $"Juntar Mesas — Salón: {_codigoSalon}";

        _flpMesas.SuspendLayout();
        foreach (var mesa in mesasSalon)
        {
            var btn = CrearBotonMesa(mesa);
            _botonesMesa.Add((btn, mesa));
            _flpMesas.Controls.Add(btn);
        }
        _flpMesas.ResumeLayout();

        ActualizarInfoSeleccion();
    }

    private Button CrearBotonMesa(Mesa mesa)
    {
        var esOriginal = mesa.CodigoMesa == _codigoMesaOriginal;
        var esSeleccionable = esOriginal || mesa.Estado == EstadoMesa.Libre;

        var btn = new Button
        {
            Text = mesa.Resumido,
            Width = 74,
            Height = 60,
            Margin = new Padding(4),
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            FlatStyle = FlatStyle.Flat,
            BackColor = esOriginal ? ColorOriginal : ColorEstado(mesa.Estado),
            ForeColor = Color.FromArgb(35, 35, 35),
            Enabled = esSeleccionable,
            Tag = mesa.CodigoMesa
        };

        btn.FlatAppearance.BorderColor = esOriginal ? ColorSeleccionBorde : ColorNormalBorde;
        btn.FlatAppearance.BorderSize = esOriginal ? 2 : 1;

        if (esSeleccionable && !esOriginal)
            btn.Click += OnMesaClick;

        return btn;
    }

    private void OnMesaClick(object? sender, EventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not string codigoMesa)
            return;

        // BR-JUNTAR-001: toggle de selección
        if (_mesasSeleccionadas.Contains(codigoMesa))
        {
            _mesasSeleccionadas.Remove(codigoMesa);
            btn.FlatAppearance.BorderColor = ColorNormalBorde;
            btn.FlatAppearance.BorderSize = 1;
        }
        else
        {
            _mesasSeleccionadas.Add(codigoMesa);
            btn.FlatAppearance.BorderColor = ColorSeleccionBorde;
            btn.FlatAppearance.BorderSize = 2;
        }

        ActualizarInfoSeleccion();
    }

    private void OnConfirmar(object? sender, EventArgs e)
    {
        // BR-JUNTAR-003: debe haber al menos la mesa original para confirmar
        if (!_mesasSeleccionadas.Contains(_codigoMesaOriginal))
        {
            MessageBox.Show("Debe seleccionar al menos la mesa original.", "Aviso",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var mesasAdicionales = _mesasSeleccionadas
            .Where(c => c != _codigoMesaOriginal)
            .OrderBy(c => c)
            .ToList();

        Resultado = new JuntarMesasResult(
            Confirmado: true,
            MesaOriginal: _codigoMesaOriginal,
            MesasAdicionales: mesasAdicionales.AsReadOnly());

        DialogResult = DialogResult.OK;
        Close();
    }

    private void ActualizarInfoSeleccion()
    {
        var seleccionadas = _mesasSeleccionadas.OrderBy(c => c).ToList();
        var nombres = _botonesMesa
            .Where(x => seleccionadas.Contains(x.Mesa.CodigoMesa))
            .Select(x => x.Mesa.Resumido)
            .ToList();

        _lblConteo.Text = seleccionadas.Count.ToString();
        _lblMesasSeleccionadas.Text = string.Join(", ", nombres);
    }

    private static Color ColorEstado(EstadoMesa estado) => estado switch
    {
        EstadoMesa.Ocupada   => ColorOcupada,
        EstadoMesa.Reservada => ColorReservada,
        EstadoMesa.EnCuenta  => Color.FromArgb(255, 165, 0),  // naranja (precuenta)
        _                    => ColorLibre
    };
}
