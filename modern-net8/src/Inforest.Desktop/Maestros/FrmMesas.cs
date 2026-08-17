using Inforest.Application.Maestros;
using Inforest.Application.Pedidos;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Mapa visual de salones y mesas — selector operativo de mesa para nuevo pedido.
/// Legacy: <c>frmMesas.frm</c>.
/// Reglas: BR-MESAS-001, BR-MESAS-002, BR-MESAS-003, BR-MESAS-004, BR-MESAS-005.
/// </summary>
public sealed class FrmMesas : Form
{
    // ── Color palette (matches legacy vbHex* constants) ──────────────────────
    private static readonly Color ClrLibre          = Color.FromArgb(0x42, 0xC6, 0x7E);  // verde
    private static readonly Color ClrOcupada        = Color.FromArgb(0xF0, 0xAD, 0x4E);  // naranja
    private static readonly Color ClrReservada      = Color.FromArgb(0x5B, 0xC0, 0xDE);  // celeste
    private static readonly Color ClrSucia          = Color.FromArgb(0xFF, 0xE0, 0x80);  // amarillo
    private static readonly Color ClrBloqueada      = Color.FromArgb(0xD9, 0x53, 0x4F);  // rojo oscuro
    private static readonly Color ClrFueraServicio  = Color.FromArgb(0xC8, 0xC8, 0xC8);  // gris claro
    private static readonly Color ClrSeleccionBorde = Color.FromArgb(0x40, 0x16, 0x0D);  // borde selección oscuro

    private static readonly Color SurfaceColor = Color.FromArgb(245, 246, 247);
    private static readonly Color AccentColor   = Color.FromArgb(0, 122, 204);
    private static readonly Font  BtnFont       = new Font("Segoe UI", 8f, FontStyle.Bold);

    // ── Dependencies ─────────────────────────────────────────────────────────
    private readonly ObtenerMesasActivosHandler      _mesasHandler;
    private readonly ObtenerSalonesActivosHandler    _salonesHandler;
    private readonly ObtenerPedidosSinMesaHandler    _sinMesaHandler;

    // ── State ────────────────────────────────────────────────────────────────
    private IReadOnlyList<Salon> _salones = [];
    private IReadOnlyList<Mesa>  _mesas   = [];
    private string? _salonActual;
    private Mesa?   _mesaSeleccionada;
    private readonly string _caja;

    // ── Result (populated on Aceptar) ────────────────────────────────────────
    public string? MesaSeleccionadaCodigo    { get; private set; }
    public string? SalonSeleccionadoCodigo   { get; private set; }
    public int NumAdultos                    { get; private set; }
    public int NumNinios                     { get; private set; }
    /// <summary>Pedido activo sin mesa seleccionado en el panel lateral.</summary>
    public string? PedidoSinMesaSeleccionado { get; private set; }

    // ── Controls ─────────────────────────────────────────────────────────────
    private readonly FlowLayoutPanel _pnlSalones   = new() { Height = 38, Dock = DockStyle.Top, BackColor = Color.FromArgb(235, 236, 240), Padding = new Padding(6, 4, 6, 4), WrapContents = false };
    private readonly FlowLayoutPanel _pnlMesas     = new() { Dock = DockStyle.Fill, Padding = new Padding(8), AutoScroll = true };
    private readonly Label           _lblInfo      = new() { Dock = DockStyle.Top, Height = 22, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0), BackColor = Color.FromArgb(235, 236, 240) };
    private readonly ListBox         _lstSinMesa   = new() { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, Font = new Font("Segoe UI", 9f) };
    private readonly NumericUpDown   _nudAdultos   = new() { Width = 60, Minimum = 0, Maximum = 999, Value = 0 };
    private readonly NumericUpDown   _nudNinios    = new() { Width = 60, Minimum = 0, Maximum = 999, Value = 0 };
    private readonly Button          _btnAceptar   = CreateButton("Aceptar",   AccentColor,  Color.White);
    private readonly Button          _btnJuntar    = CreateButton("Juntar",    Color.FromArgb(100,100,100), Color.White);
    private readonly Button          _btnActualizar = CreateButton("Actualizar", Color.FromArgb(70,130,180), Color.White);
    private readonly Button          _btnVolver    = CreateButton("Volver",    Color.FromArgb(108, 117, 125), Color.White);
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 30_000 };

    public FrmMesas(
        ObtenerMesasActivosHandler mesasHandler,
        ObtenerSalonesActivosHandler salonesHandler,
        ObtenerPedidosSinMesaHandler sinMesaHandler,
        string caja = "")
    {
        _mesasHandler   = mesasHandler;
        _salonesHandler = salonesHandler;
        _sinMesaHandler = sinMesaHandler;
        _caja           = caja;

        Text            = "Mapa de Mesas";
        Width           = 1200;
        Height          = 720;
        StartPosition   = FormStartPosition.CenterScreen;
        BackColor       = SurfaceColor;
        MinimizeBox     = false;
        MaximizeBox     = false;
        FormBorderStyle = FormBorderStyle.Sizable;

        BuildLayout();

        Load             += async (_, _) => await CargarAsync();
        _btnAceptar.Click  += BtnAceptar_Click;
        _btnActualizar.Click += async (_, _) => await CargarAsync();
        _btnVolver.Click   += (_, _) => Close();
        _btnJuntar.Click   += BtnJuntar_Click;
        _lstSinMesa.DoubleClick += LstSinMesa_DoubleClick;
        _timer.Tick        += async (_, _) => await CargarAsync();
    }

    // ── Layout ────────────────────────────────────────────────────────────────
    private void BuildLayout()
    {
        var main = new TableLayoutPanel
        {
            Dock        = DockStyle.Fill,
            RowCount    = 3,
            ColumnCount = 2
        };
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 75));
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));  // salones
        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));  // info
        main.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // mesas

        main.Controls.Add(_pnlSalones, 0, 0);
        main.SetColumnSpan(_pnlSalones, 2);

        main.Controls.Add(_lblInfo, 0, 1);
        main.SetColumnSpan(_lblInfo, 2);

        main.Controls.Add(_pnlMesas, 0, 2);

        // Right panel: pedidos sin mesa + controls
        var pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(4), BackColor = Color.FromArgb(240, 241, 245) };

        var lblSinMesa = new Label { Text = "Pedidos sin mesa", Dock = DockStyle.Top, Height = 20, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), ForeColor = Color.FromArgb(60, 60, 60) };

        // Guests
        var pnlGuests = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, AutoSize = false };
        pnlGuests.Controls.Add(new Label { Text = "Adultos:", AutoSize = true, Margin = new Padding(0, 8, 4, 0) });
        pnlGuests.Controls.Add(_nudAdultos);
        pnlGuests.Controls.Add(new Label { Text = "Niños:", AutoSize = true, Margin = new Padding(8, 8, 4, 0) });
        pnlGuests.Controls.Add(_nudNinios);

        // Action buttons
        var pnlBtns = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, FlowDirection = FlowDirection.LeftToRight };
        pnlBtns.Controls.AddRange([_btnAceptar, _btnJuntar, _btnActualizar, _btnVolver]);

        pnlRight.Controls.Add(_lstSinMesa);
        pnlRight.Controls.Add(pnlGuests);
        pnlRight.Controls.Add(lblSinMesa);
        pnlRight.Controls.Add(pnlBtns);

        main.Controls.Add(pnlRight, 1, 2);

        // Legend
        var legend = BuildLegend();
        legend.Dock = DockStyle.Bottom;
        legend.Height = 28;

        Controls.Add(main);
        Controls.Add(legend);
    }

    private static Panel BuildLegend()
    {
        var p = new FlowLayoutPanel { BackColor = Color.FromArgb(235, 236, 240), Padding = new Padding(8, 4, 8, 4), WrapContents = false };
        void Add(Color c, string label)
        {
            var swatch = new Panel { Width = 16, Height = 16, BackColor = c, Margin = new Padding(2, 2, 0, 0), BorderStyle = BorderStyle.FixedSingle };
            var lbl    = new Label  { Text = label, AutoSize = true, Margin = new Padding(2, 2, 10, 0) };
            p.Controls.Add(swatch);
            p.Controls.Add(lbl);
        }
        Add(ClrLibre,          "Libre");
        Add(ClrOcupada,        "Ocupada");
        Add(ClrReservada,      "Reservada");
        Add(ClrSucia,          "Sucia");
        Add(ClrBloqueada,      "Bloqueada");
        Add(ClrFueraServicio,  "Fuera de servicio");
        return p;
    }

    // ── Data loading ──────────────────────────────────────────────────────────
    private async Task CargarAsync()
    {
        _timer.Stop();
        try
        {
            var rSalones = await _salonesHandler.HandleAsync(new ObtenerSalonesActivosQuery());
            var rMesas   = await _mesasHandler.HandleAsync(new ObtenerMesasActivosQuery());

            if (!rSalones.EsExitoso || !rMesas.EsExitoso) return;

            _salones = rSalones.Valor!;
            _mesas   = rMesas.Valor!;

            RenderizarSalones();
            if (_salonActual == null && _salones.Count > 0)
                SeleccionarSalon(_salones[0].CodigoSalon);
            else
                RenderizarMesas();

            await CargarSinMesaAsync();
        }
        finally
        {
            _timer.Start();
        }
    }

    private async Task CargarSinMesaAsync()
    {
        if (string.IsNullOrEmpty(_caja)) return;
        var r = await _sinMesaHandler.HandleAsync(new ObtenerPedidosSinMesaQuery(_caja));
        if (!r.EsExitoso) return;
        _lstSinMesa.Items.Clear();
        foreach (var p in r.Valor!)
            _lstSinMesa.Items.Add(new PedidoSinMesaItem(p.CodigoPedido, p.Observacion));
    }

    // ── Render salones ────────────────────────────────────────────────────────
    private void RenderizarSalones()
    {
        _pnlSalones.Controls.Clear();
        foreach (var s in _salones)
        {
            var btn = new Button
            {
                Text      = s.Resumido,
                Tag       = s.CodigoSalon,
                Height    = 28,
                AutoSize  = true,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 8.5f),
                BackColor = s.CodigoSalon == _salonActual ? AccentColor : Color.FromArgb(255, 255, 255),
                ForeColor = s.CodigoSalon == _salonActual ? Color.White : Color.FromArgb(40, 40, 40),
                Margin    = new Padding(2, 2, 2, 2)
            };
            btn.FlatAppearance.BorderColor = s.CodigoSalon == _salonActual ? AccentColor : Color.FromArgb(200, 200, 200);
            btn.Click += (_, _) => SeleccionarSalon((string)btn.Tag!);
            _pnlSalones.Controls.Add(btn);
        }
    }

    private void SeleccionarSalon(string codigoSalon)
    {
        _salonActual       = codigoSalon;
        _mesaSeleccionada  = null;
        RenderizarSalones();
        RenderizarMesas();
    }

    // ── Render mesas ──────────────────────────────────────────────────────────
    private void RenderizarMesas()
    {
        _pnlMesas.SuspendLayout();
        _pnlMesas.Controls.Clear();

        var mesasSalon = _mesas
            .Where(m => m.CodigoSalon == _salonActual && m.Activo)
            .OrderBy(m => m.CodigoMesa)
            .ToList();

        int libre = mesasSalon.Count(m => m.Estado is EstadoMesa.Libre or EstadoMesa.Sucia);
        var salon = _salones.FirstOrDefault(s => s.CodigoSalon == _salonActual);
        _lblInfo.Text = $"{salon?.Detallado ?? _salonActual}   {libre}/{mesasSalon.Count} mesa(s) disponible(s)";

        foreach (var mesa in mesasSalon)
        {
            var btn = new Panel
            {
                Width     = 80,
                Height    = 60,
                Margin    = new Padding(4),
                BackColor = ColorPorEstado(mesa.Estado),
                BorderStyle = BorderStyle.FixedSingle,
                Cursor    = EsSeleccionable(mesa) ? Cursors.Hand : Cursors.Default,
                Tag       = mesa
            };

            var lblResumido = new Label
            {
                Text      = mesa.Resumido,
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font      = BtnFont,
                BackColor = Color.Transparent,
                ForeColor = mesa.Estado == EstadoMesa.FueraDeServicio ? Color.Gray : Color.FromArgb(35, 35, 35)
            };
            btn.Controls.Add(lblResumido);

            if (mesa == _mesaSeleccionada)
                btn.BackColor = ControlPaint.Dark(ColorPorEstado(mesa.Estado), 0.15f);

            btn.Click     += (_, _) => MesaClick(mesa, btn);
            lblResumido.Click += (_, _) => MesaClick(mesa, btn);

            _pnlMesas.Controls.Add(btn);
        }

        _pnlMesas.ResumeLayout();
    }

    private void MesaClick(Mesa mesa, Panel btn)
    {
        // BR-MESAS-002: mesa ocupada → abrir detalle del pedido
        if (mesa.Estado == EstadoMesa.Ocupada)
        {
            // Notifica al llamador para que abra FrmDetallePedido con el pedido de esa mesa.
            MesaOcupadaSeleccionada?.Invoke(this, mesa.CodigoMesa);
            return;
        }

        // BR-MESAS-001: libre o sucia → seleccionar
        if (mesa.Estado is EstadoMesa.Libre or EstadoMesa.Sucia)
        {
            if (_mesaSeleccionada == mesa)
            {
                _mesaSeleccionada = null;
            }
            else
            {
                _mesaSeleccionada = mesa;
            }
            RenderizarMesas();
        }
        // BR-MESAS-003: reservada/bloqueada/fuera de servicio → sin acción
    }

    // ── Buttons ───────────────────────────────────────────────────────────────
    private void BtnAceptar_Click(object? sender, EventArgs e)
    {
        // BR-MESAS-004: retornar mesa + contadores al llamador
        if (_mesaSeleccionada == null)
        {
            MessageBox.Show("Seleccione una mesa disponible para continuar.", "Sin selección",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var res = MessageBox.Show(
            $"¿Crear pedido con mesa {_mesaSeleccionada.CodigoMesa} — {_mesaSeleccionada.Resumido}?",
            $"Mesa {_mesaSeleccionada.CodigoMesa}",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (res != DialogResult.Yes) return;

        MesaSeleccionadaCodigo  = _mesaSeleccionada.CodigoMesa;
        SalonSeleccionadoCodigo = _mesaSeleccionada.CodigoSalon;
        NumAdultos              = (int)_nudAdultos.Value;
        NumNinios               = (int)_nudNinios.Value;
        DialogResult            = DialogResult.OK;
        Close();
    }

    private void BtnJuntar_Click(object? sender, EventArgs e)
    {
        if (_mesaSeleccionada == null)
        {
            MessageBox.Show("Seleccione una mesa base para juntar.", "Sin selección",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        // Notifica al llamador: el caller abrirá FrmJuntarMesas con la mesa seleccionada como base.
        JuntarMesasSolicitado?.Invoke(this, _mesaSeleccionada.CodigoMesa);
    }

    private void LstSinMesa_DoubleClick(object? sender, EventArgs e)
    {
        if (_lstSinMesa.SelectedItem is PedidoSinMesaItem item)
        {
            PedidoSinMesaSeleccionado = item.CodigoPedido;
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    // ── Events for parent coordination ───────────────────────────────────────
    /// <summary>Fired when a <see cref="EstadoMesa.Ocupada"/> table is clicked; arg = CodigoMesa.</summary>
    public event EventHandler<string>? MesaOcupadaSeleccionada;
    /// <summary>Fired when "Juntar" is requested with a selected table; arg = CodigoMesa.</summary>
    public event EventHandler<string>? JuntarMesasSolicitado;

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static bool EsSeleccionable(Mesa m)
        => m.Estado is EstadoMesa.Libre or EstadoMesa.Sucia or EstadoMesa.Ocupada;

    private static Color ColorPorEstado(EstadoMesa estado) => estado switch
    {
        EstadoMesa.Libre         => ClrLibre,
        EstadoMesa.Ocupada       => ClrOcupada,
        EstadoMesa.Reservada     => ClrReservada,
        EstadoMesa.Sucia         => ClrSucia,
        EstadoMesa.Bloqueada     => ClrBloqueada,
        EstadoMesa.FueraDeServicio => ClrFueraServicio,
        EstadoMesa.EnCuenta      => ClrOcupada,
        _                        => ClrFueraServicio
    };

    private static Button CreateButton(string text, Color back, Color fore) => new()
    {
        Text      = text,
        Height    = 32,
        Width     = 90,
        FlatStyle = FlatStyle.Flat,
        BackColor = back,
        ForeColor = fore,
        Font      = new Font("Segoe UI", 8.5f),
        Margin    = new Padding(2)
    };

    // ── Override dispose ──────────────────────────────────────────────────────
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _timer.Dispose();
            BtnFont.Dispose();
        }
        base.Dispose(disposing);
    }
}

/// <summary>Item de ListBox para pedidos sin mesa.</summary>
file sealed record PedidoSinMesaItem(string CodigoPedido, string Observacion)
{
    public override string ToString()
        => string.IsNullOrWhiteSpace(Observacion)
            ? CodigoPedido
            : $"{CodigoPedido} — {Observacion}";
}
