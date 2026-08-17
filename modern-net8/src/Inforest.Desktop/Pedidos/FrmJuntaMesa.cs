using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Pedidos;

/// <summary>
/// Gestión visual de junta de mesas para un pedido.
/// Legacy: frmJuntaMesa.frm — asigna múltiples mesas físicas a un mismo pedido
/// mediante la tabla TPEDIDOMESA.
/// BR-JUNTA-001: limpia asignación anterior antes de insertar nuevas mesas.
/// BR-JUNTA-002: solo mesas Libres (01) o Sucias (04) pueden ser juntadas.
/// BR-JUNTA-003: las mesas juntas adoptan estado '06' (FueraDeServicio) en legacy.
/// </summary>
public sealed class FrmJuntaMesa : Form
{
    // ── Dependencias ─────────────────────────────────────────────────────────
    private readonly ObtenerMesasActivosHandler     _mesasHandler;
    private readonly ObtenerMesasJuntadasHandler    _juntadasHandler;
    private readonly ActualizarJuntaMesasHandler    _actualizarHandler;
    private readonly string                         _codigoPedido;
    private readonly string                         _codigoSalon;

    // ── Estado ───────────────────────────────────────────────────────────────
    private readonly HashSet<string>   _mesasSeleccionadas = [];
    private List<Mesa>                 _mesas = [];
    private readonly List<Button>      _botones = [];

    // ── Controles ─────────────────────────────────────────────────────────────
    private readonly Label           _lblConteo;
    private readonly Label           _lblStatus;
    private readonly FlowLayoutPanel _panelMesas;
    private readonly Button          _btnAceptar;

    // Colores
    private static readonly Color ColorLibre    = Color.FromArgb( 76, 175,  80);
    private static readonly Color ColorOcupada  = Color.FromArgb(244,  67,  54);
    private static readonly Color ColorReserv   = Color.FromArgb(255, 152,   0);
    private static readonly Color ColorSucia    = Color.FromArgb(121,  85, 172);
    private static readonly Color ColorBloquead = Color.FromArgb(100, 100, 100);
    private static readonly Color ColorJunta    = Color.FromArgb( 33, 150, 243);  // '06' juntas

    public FrmJuntaMesa(
        ObtenerMesasActivosHandler  mesasHandler,
        ObtenerMesasJuntadasHandler juntadasHandler,
        ActualizarJuntaMesasHandler actualizarHandler,
        string codigoPedido,
        string codigoSalon)
    {
        _mesasHandler     = mesasHandler;
        _juntadasHandler  = juntadasHandler;
        _actualizarHandler = actualizarHandler;
        _codigoPedido     = codigoPedido;
        _codigoSalon      = codigoSalon;

        Text            = "Junta de Mesas";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition   = FormStartPosition.CenterParent;
        MaximizeBox     = false;
        MinimizeBox     = false;
        Width           = 700;
        Height          = 540;
        BackColor       = Color.FromArgb(245, 246, 247);

        // ── Header ──────────────────────────────────────────────────────────
        var header = new Panel
        {
            Dock      = DockStyle.Top,
            Height    = 40,
            BackColor = Color.FromArgb(40, 60, 90)
        };
        var lblTitulo = new Label
        {
            Text      = "Seleccione las mesas a juntar al pedido",
            Dock      = DockStyle.Fill,
            Font      = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.White,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding   = new Padding(12, 0, 0, 0)
        };
        header.Controls.Add(lblTitulo);

        // ── Leyenda ──────────────────────────────────────────────────────────
        var leyenda = new Panel
        {
            Dock      = DockStyle.Top,
            Height    = 32,
            BackColor = Color.FromArgb(235, 236, 240)
        };
        var flowLeyenda = new FlowLayoutPanel
        {
            Dock          = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            Padding       = new Padding(8, 4, 8, 4)
        };
        void AgregarChip(string texto, Color color)
        {
            flowLeyenda.Controls.Add(new Label
            {
                Text      = $"  {texto}  ",
                AutoSize  = true,
                BackColor = color,
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 7.5f),
                Padding   = new Padding(3, 1, 3, 1),
                Margin    = new Padding(3, 2, 3, 2)
            });
        }
        AgregarChip("Libre (elegible)", ColorLibre);
        AgregarChip("Sucia (elegible)", ColorSucia);
        AgregarChip("Ocupada", ColorOcupada);
        AgregarChip("Juntada/Seleccionada", ColorJunta);
        leyenda.Controls.Add(flowLeyenda);

        // ── Mesas ────────────────────────────────────────────────────────────
        _panelMesas = new FlowLayoutPanel
        {
            Dock          = DockStyle.Fill,
            AutoScroll    = true,
            Padding       = new Padding(8),
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents  = true
        };

        // ── Footer ───────────────────────────────────────────────────────────
        _lblConteo = new Label
        {
            Text      = "0 mesas juntas",
            AutoSize  = true,
            Padding   = new Padding(6, 0, 0, 0),
            Font      = new Font("Segoe UI", 9)
        };
        _lblStatus = new Label
        {
            Text      = "",
            AutoSize  = true,
            ForeColor = Color.DarkRed,
            Padding   = new Padding(6, 0, 0, 0),
            Font      = new Font("Segoe UI", 8.5f)
        };

        var statusPanel = new FlowLayoutPanel
        {
            Dock          = DockStyle.Left,
            Width         = 360,
            FlowDirection = FlowDirection.TopDown,
            Padding       = new Padding(4, 8, 0, 0)
        };
        statusPanel.Controls.AddRange([_lblConteo, _lblStatus]);

        _btnAceptar = new Button
        {
            Text      = "Guardar",
            Width     = 100,
            Height    = 36,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(40, 130, 70),
            ForeColor = Color.White,
            Font      = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor    = Cursors.Hand,
            Margin    = new Padding(4, 0, 0, 0)
        };
        _btnAceptar.FlatAppearance.BorderSize = 0;
        _btnAceptar.Click += async (_, _) => await GuardarAsync();

        var btnCancelar = new Button
        {
            Text      = "Cancelar",
            Width     = 100,
            Height    = 36,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            Font      = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor    = Cursors.Hand,
            Margin    = new Padding(4, 0, 0, 0),
            DialogResult = DialogResult.Cancel
        };
        btnCancelar.FlatAppearance.BorderSize = 0;
        btnCancelar.Click += (_, _) => Close();

        var footer = new FlowLayoutPanel
        {
            Dock          = DockStyle.Bottom,
            Height        = 52,
            FlowDirection = FlowDirection.RightToLeft,
            Padding       = new Padding(8, 6, 8, 6),
            BackColor     = Color.FromArgb(235, 236, 240)
        };
        footer.Controls.AddRange([btnCancelar, _btnAceptar, statusPanel]);

        Controls.Add(_panelMesas);
        Controls.Add(footer);
        Controls.Add(leyenda);
        Controls.Add(header);

        Load += async (_, _) => await CargarAsync();
    }

    // ── Carga ─────────────────────────────────────────────────────────────────

    private async Task CargarAsync()
    {
        _lblStatus.Text = "Cargando...";
        _btnAceptar.Enabled = false;

        var rMesas    = await _mesasHandler.HandleAsync(new ObtenerMesasActivosQuery());
        var rJuntadas = await _juntadasHandler.HandleAsync(new ObtenerMesasJuntadasQuery(_codigoPedido));

        if (!rMesas.EsExitoso) { _lblStatus.Text = "Error al cargar mesas."; return; }

        _mesas = rMesas.Valor!.Where(m => m.CodigoSalon == _codigoSalon).ToList();

        if (rJuntadas.EsExitoso)
            foreach (var c in rJuntadas.Valor!) _mesasSeleccionadas.Add(c);

        RenderizarMesas();
        ActualizarConteo();
        _btnAceptar.Enabled = true;
        _lblStatus.Text = "";
    }

    private void RenderizarMesas()
    {
        _panelMesas.SuspendLayout();
        _panelMesas.Controls.Clear();
        _botones.Clear();

        foreach (var mesa in _mesas)
        {
            var esSeleccionada = _mesasSeleccionadas.Contains(mesa.CodigoMesa);
            var esElegible     = mesa.Estado is EstadoMesa.Libre or EstadoMesa.Sucia or EstadoMesa.FueraDeServicio;

            var color = esSeleccionada ? ColorJunta
                : mesa.Estado switch
                {
                    EstadoMesa.Libre           => ColorLibre,
                    EstadoMesa.Ocupada         => ColorOcupada,
                    EstadoMesa.Reservada       => ColorReserv,
                    EstadoMesa.Sucia           => ColorSucia,
                    EstadoMesa.Bloqueada       => ColorBloquead,
                    EstadoMesa.FueraDeServicio => ColorJunta,
                    _ => Color.Gray
                };

            var btn = new Button
            {
                Width     = 88,
                Height    = 76,
                Text      = mesa.Resumido,
                FlatStyle = FlatStyle.Flat,
                BackColor = color,
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor    = esElegible ? Cursors.Hand : Cursors.No,
                Enabled   = esElegible,
                Margin    = new Padding(4),
                Tag       = mesa
            };
            btn.FlatAppearance.BorderSize = 1;
            btn.Click += (_, _) => ToggleMesa(btn, mesa);
            _panelMesas.Controls.Add(btn);
            _botones.Add(btn);
        }

        _panelMesas.ResumeLayout();
    }

    private void ToggleMesa(Button btn, Mesa mesa)
    {
        if (_mesasSeleccionadas.Contains(mesa.CodigoMesa))
        {
            _mesasSeleccionadas.Remove(mesa.CodigoMesa);
            btn.BackColor = mesa.Estado switch
            {
                EstadoMesa.Libre  => ColorLibre,
                EstadoMesa.Sucia  => ColorSucia,
                _ => ColorJunta
            };
        }
        else
        {
            _mesasSeleccionadas.Add(mesa.CodigoMesa);
            btn.BackColor = ColorJunta;
        }
        ActualizarConteo();
    }

    private void ActualizarConteo()
        => _lblConteo.Text = $"{_mesasSeleccionadas.Count} mesa(s) seleccionada(s)";

    private async Task GuardarAsync()
    {
        _btnAceptar.Enabled = false;
        _lblStatus.Text     = "Guardando...";

        var r = await _actualizarHandler.HandleAsync(
            new ActualizarJuntaMesasCommand(_codigoPedido, _mesasSeleccionadas));

        if (!r.EsExitoso)
        {
            _lblStatus.Text     = $"Error: {r.MensajeError}";
            _btnAceptar.Enabled = true;
            return;
        }

        DialogResult = DialogResult.OK;
        Close();
    }
}
