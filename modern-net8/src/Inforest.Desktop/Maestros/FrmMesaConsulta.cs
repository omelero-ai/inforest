using Inforest.Application.Maestros;
using Inforest.Application.Pedidos;
using Inforest.Domain.Entities.Maestros;
using Inforest.Desktop.Shared;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Mapa visual de mesas con soporte para tres modos operativos.
/// Legacy: frmMesaConsulta.frm.
/// BR-MESACONSULTA-001..004.
/// </summary>
public enum ModoConsulta
{
    /// <summary>Visualización de estados y cambio de estado no-ocupadas (sTipo="V" en VB6).</summary>
    Visual,
    /// <summary>Selección de mesa para nuevo pedido (default en VB6).</summary>
    Seleccionar,
    /// <summary>Selección de mesa destino para transferencia de pedido (sTipo="M" en VB6).</summary>
    Mover
}

/// <summary>
/// Mapa visual de mesas con soporte para tres modos operativos.
/// Legacy: frmMesaConsulta.frm.
/// Modos: <see cref="ModoConsulta.Visual"/> (ver/cambiar estado), <see cref="ModoConsulta.Seleccionar"/> (asignar mesa a pedido),
/// <see cref="ModoConsulta.Mover"/> (transferir pedido a otra mesa).
/// </summary>
public sealed class FrmMesaConsulta : Form
{
    // ── Dependencias ─────────────────────────────────────────────────────────
    private readonly ObtenerMesasActivosHandler          _mesasHandler;
    private readonly ObtenerSalonesActivosHandler        _salonesHandler;
    private readonly CambiarEstadoMesaHandler            _cambiarEstadoHandler;
    private readonly ObtenerPedidosSinMesaHandler?       _sinMesaHandler;
    private readonly ModoConsulta                        _modo;

    // ── Resultado ─────────────────────────────────────────────────────────────
    /// <summary>Mesa seleccionada por el usuario. Null si canceló o eligió "Sin Mesa".</summary>
    public Mesa?   MesaSeleccionada   { get; private set; }

    /// <summary>True si el usuario eligió "Sin Mesa" (wMesa=True, sCodigo="" en VB6).</summary>
    public bool    EsSinMesa          { get; private set; }

    /// <summary>Número de adultos ingresado (0 si no aplica al modo).</summary>
    public int     NumAdultos         { get; private set; }

    /// <summary>Número de niños ingresado (0 si no aplica al modo).</summary>
    public int     NumNinios          { get; private set; }

    // ── Estado interno ────────────────────────────────────────────────────────
    private List<Mesa>   _mesas   = [];
    private List<Salon>  _salones = [];
    private int          _salonActualIdx;
    private string?      _codigoCaja;

    // ── Controles ─────────────────────────────────────────────────────────────
    private readonly FlowLayoutPanel _panelSalones;
    private readonly FlowLayoutPanel _panelMesas;
    private readonly Label           _lblSalonNombre;
    private readonly Label           _lblStatus;
    private readonly Panel           _panelEstados;
    private readonly Panel           _panelContadores;
    private readonly ListBox         _lstSinMesa;

    // Contadores (visibles en modo Seleccionar/Mover)
    private readonly NumericUpDown _nudAdultos = new() { Width = 60, Minimum = 0, Maximum = 999, Value = 0 };
    private readonly NumericUpDown _nudNinios  = new() { Width = 60, Minimum = 0, Maximum = 999, Value = 0 };

    // Botones de acción
    private readonly Button _btnAceptar;
    private readonly Button _btnSinMesa;
    private readonly Button _btnRefrescar;
    private readonly Button _btnSalir;

    // ── Colores por estado ────────────────────────────────────────────────────
    private static readonly Dictionary<EstadoMesa, Color> ColoresMesa = new()
    {
        { EstadoMesa.Libre,           Color.FromArgb( 76, 175,  80) },
        { EstadoMesa.Ocupada,         Color.FromArgb(244,  67,  54) },
        { EstadoMesa.Reservada,       Color.FromArgb(255, 152,   0) },
        { EstadoMesa.Sucia,           Color.FromArgb(121,  85, 172) },
        { EstadoMesa.Bloqueada,       Color.FromArgb(100, 100, 100) },
        { EstadoMesa.FueraDeServicio, Color.FromArgb( 55,  71,  79) },
        { EstadoMesa.EnCuenta,        Color.FromArgb( 30, 136, 229) }
    };

    private static readonly Dictionary<EstadoMesa, string> LabelEstado = new()
    {
        { EstadoMesa.Libre,           "Libre"          },
        { EstadoMesa.Ocupada,         "Ocupada"        },
        { EstadoMesa.Reservada,       "Reservada"      },
        { EstadoMesa.Sucia,           "Sucia"          },
        { EstadoMesa.Bloqueada,       "Bloqueada"      },
        { EstadoMesa.FueraDeServicio, "Fuera Servicio" },
        { EstadoMesa.EnCuenta,        "En Cuenta"      }
    };

    /// <summary>
    /// Crea el mapa de consulta de mesas.
    /// </summary>
    /// <param name="mesasHandler">Handler de consulta de mesas.</param>
    /// <param name="salonesHandler">Handler de consulta de salones.</param>
    /// <param name="cambiarEstadoHandler">Handler para cambiar estado (requerido en modo Visual).</param>
    /// <param name="sinMesaHandler">Handler para pedidos sin mesa (opcional, usado en modo Seleccionar).</param>
    /// <param name="modo">Modo de operación: Visual, Seleccionar o Mover.</param>
    /// <param name="codigoCaja">Código de caja (para filtrar pedidos sin mesa).</param>
    public FrmMesaConsulta(
        ObtenerMesasActivosHandler   mesasHandler,
        ObtenerSalonesActivosHandler salonesHandler,
        CambiarEstadoMesaHandler     cambiarEstadoHandler,
        ObtenerPedidosSinMesaHandler? sinMesaHandler = null,
        ModoConsulta                 modo            = ModoConsulta.Seleccionar,
        string?                      codigoCaja      = null)
    {
        _mesasHandler         = mesasHandler;
        _salonesHandler       = salonesHandler;
        _cambiarEstadoHandler = cambiarEstadoHandler;
        _sinMesaHandler       = sinMesaHandler;
        _modo                 = modo;
        _codigoCaja           = codigoCaja;

        string titulo = modo switch
        {
            ModoConsulta.Visual      => "Mesas — Consulta y Estado",
            ModoConsulta.Mover       => "Mesas — Seleccionar Mesa Destino",
            _                        => "Mesas — Seleccionar Mesa"
        };

        Text            = titulo;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition   = FormStartPosition.CenterParent;
        MaximizeBox     = false;
        MinimizeBox     = false;
        Width           = 960;
        Height          = 680;
        BackColor       = Color.FromArgb(245, 246, 247);

        // ── Header ──────────────────────────────────────────────────────────
        var header = new Panel
        {
            Dock      = DockStyle.Top,
            Height    = 42,
            BackColor = Color.FromArgb(40, 60, 90)
        };
        _lblSalonNombre = new Label
        {
            Text      = "",
            Dock      = DockStyle.Fill,
            Font      = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.White,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding   = new Padding(12, 0, 0, 0)
        };
        header.Controls.Add(_lblSalonNombre);

        // ── Salones (tab buttons) ────────────────────────────────────────────
        _panelSalones = new FlowLayoutPanel
        {
            Dock          = DockStyle.Top,
            Height        = 48,
            Padding       = new Padding(6, 4, 6, 4),
            FlowDirection = FlowDirection.LeftToRight,
            BackColor     = Color.FromArgb(30, 50, 80)
        };

        // ── Panel de estados (leyenda, visible en modo Visual) ───────────────
        _panelEstados = new Panel
        {
            Dock      = DockStyle.Bottom,
            Height    = modo == ModoConsulta.Visual ? 52 : 0,
            BackColor = Color.FromArgb(235, 236, 240),
            Visible   = modo == ModoConsulta.Visual
        };
        if (modo == ModoConsulta.Visual)
            CrearLeyendaEstados();

        // ── Panel de contadores (adultos/niños) ──────────────────────────────
        _panelContadores = new Panel
        {
            Dock      = DockStyle.Bottom,
            Height    = (modo != ModoConsulta.Visual) ? 44 : 0,
            BackColor = Color.FromArgb(235, 236, 240),
            Visible   = modo != ModoConsulta.Visual
        };
        if (modo != ModoConsulta.Visual)
            CrearContadores();

        // ── Lista sin mesa ───────────────────────────────────────────────────
        _lstSinMesa = new ListBox
        {
            Dock      = DockStyle.Right,
            Width     = 180,
            Font      = new Font("Segoe UI", 8.5f),
            Visible   = modo == ModoConsulta.Seleccionar
        };
        var lblSinMesa = new Label
        {
            Text      = "Sin mesa",
            Dock      = DockStyle.Top,
            Height    = 22,
            Font      = new Font("Segoe UI", 8, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleCenter,
            BackColor = Color.FromArgb(70, 90, 120),
            ForeColor = Color.White,
            Visible   = modo == ModoConsulta.Seleccionar
        };
        var panelDerecho = new Panel { Dock = DockStyle.Right, Width = 182 };
        panelDerecho.Controls.AddRange([lblSinMesa, _lstSinMesa]);

        // ── Mesas grid ───────────────────────────────────────────────────────
        _panelMesas = new FlowLayoutPanel
        {
            Dock          = DockStyle.Fill,
            AutoScroll    = true,
            Padding       = new Padding(8),
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents  = true
        };

        // ── Footer ───────────────────────────────────────────────────────────
        var footer = new FlowLayoutPanel
        {
            Dock          = DockStyle.Bottom,
            Height        = 52,
            Padding       = new Padding(8, 6, 8, 6),
            FlowDirection = FlowDirection.RightToLeft,
            BackColor     = Color.FromArgb(235, 236, 240)
        };

        _btnSalir = CrearBotonAccion("Salir", Color.FromArgb(220, 53, 69));
        _btnSalir.Click += (_, _) => { MesaSeleccionada = null; DialogResult = DialogResult.Cancel; Close(); };

        _btnRefrescar = CrearBotonAccion("Actualizar", Color.FromArgb(70, 100, 150));
        _btnRefrescar.Click += async (_, _) => await CargarAsync();

        _btnAceptar = CrearBotonAccion("Aceptar", Color.FromArgb(40, 130, 70));
        _btnAceptar.Enabled = false;
        _btnAceptar.Click += (_, _) =>
        {
            if (modo != ModoConsulta.Visual)
            {
                NumAdultos = (int)_nudAdultos.Value;
                NumNinios  = (int)_nudNinios.Value;
            }
            DialogResult = DialogResult.OK;
            Close();
        };

        _btnSinMesa = CrearBotonAccion("Sin Mesa", Color.FromArgb(100, 120, 160));
        _btnSinMesa.Visible = modo == ModoConsulta.Seleccionar;
        _btnSinMesa.Click += (_, _) =>
        {
            EsSinMesa  = true;
            NumAdultos = (int)_nudAdultos.Value;
            NumNinios  = (int)_nudNinios.Value;
            DialogResult = DialogResult.OK;
            Close();
        };

        _lblStatus = new Label
        {
            AutoSize  = true,
            Text      = "Cargando...",
            Padding   = new Padding(4, 0, 0, 0),
            Anchor    = AnchorStyles.Left | AnchorStyles.Top
        };
        var statusWrap = new FlowLayoutPanel
        {
            Dock          = DockStyle.Left,
            Width         = 300,
            FlowDirection = FlowDirection.LeftToRight,
            Padding       = new Padding(4, 8, 0, 0)
        };
        statusWrap.Controls.Add(_lblStatus);

        footer.Controls.AddRange([_btnSalir, _btnRefrescar, _btnAceptar, _btnSinMesa]);
        footer.Controls.Add(statusWrap);

        // ── Ensamble ─────────────────────────────────────────────────────────
        Controls.Add(_panelMesas);
        Controls.Add(panelDerecho);
        Controls.Add(_panelContadores);
        Controls.Add(_panelEstados);
        Controls.Add(footer);
        Controls.Add(_panelSalones);
        Controls.Add(header);

        Load += async (_, _) => await CargarAsync();
    }

    // ── Carga ─────────────────────────────────────────────────────────────────

    private async Task CargarAsync()
    {
        _btnAceptar.Enabled   = false;
        _btnRefrescar.Enabled = false;
        _lblStatus.Text       = "Cargando mesas...";

        try
        {
            var rMesas   = await _mesasHandler.HandleAsync(new ObtenerMesasActivosQuery());
            var rSalones = await _salonesHandler.HandleAsync(new ObtenerSalonesActivosQuery());

            if (!rMesas.EsExitoso || !rSalones.EsExitoso)
            {
                _lblStatus.Text = "Error al cargar datos.";
                return;
            }

            _mesas   = [.. rMesas.Valor!];
            _salones = [.. rSalones.Valor!];

            RenderizarSalones();
            if (_salones.Count > 0) MostrarSalon(0);

            if (_modo == ModoConsulta.Seleccionar && _sinMesaHandler is not null && _codigoCaja is not null)
                await CargarSinMesaAsync();

            _lblStatus.Text = $"{_mesas.Count} mesas activas.";
        }
        catch (Exception ex)
        {
            _lblStatus.Text = $"Error: {ex.Message}";
        }
        finally
        {
            _btnRefrescar.Enabled = true;
        }
    }

    private async Task CargarSinMesaAsync()
    {
        if (_sinMesaHandler is null || _codigoCaja is null) return;
        var r = await _sinMesaHandler.HandleAsync(new ObtenerPedidosSinMesaQuery(_codigoCaja));
        if (!r.EsExitoso) return;
        _lstSinMesa.Items.Clear();
        foreach (var p in r.Valor!)
            _lstSinMesa.Items.Add(p.Observacion.Length > 0 ? p.Observacion : p.CodigoPedido);
    }

    // ── Renderizado ──────────────────────────────────────────────────────────

    private void RenderizarSalones()
    {
        _panelSalones.Controls.Clear();
        for (int i = 0; i < _salones.Count; i++)
        {
            int idx = i;
            var salon = _salones[i];
            var btn = new Button
            {
                Text      = salon.Detallado,
                Height    = 34,
                Width     = Math.Max(80, salon.Detallado.Length * 9),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 90, 130),
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                Cursor    = Cursors.Hand,
                Margin    = new Padding(2, 0, 2, 0),
                Tag       = idx
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (_, _) => MostrarSalon(idx);
            _panelSalones.Controls.Add(btn);
        }
    }

    private void MostrarSalon(int idx)
    {
        if (idx < 0 || idx >= _salones.Count) return;
        _salonActualIdx = idx;
        var salon = _salones[idx];
        _lblSalonNombre.Text = $"Salón: {salon.Detallado}";

        // Marcar botón activo
        for (int i = 0; i < _panelSalones.Controls.Count; i++)
        {
            if (_panelSalones.Controls[i] is Button b)
                b.BackColor = (int)(b.Tag ?? -1) == idx
                    ? Color.FromArgb(40, 160, 90)
                    : Color.FromArgb(70, 90, 130);
        }

        var mesasSalon = _mesas.Where(m => m.CodigoSalon == salon.CodigoSalon).ToList();
        RenderizarMesas(mesasSalon);
    }

    private void RenderizarMesas(List<Mesa> mesas)
    {
        _panelMesas.SuspendLayout();
        _panelMesas.Controls.Clear();

        foreach (var mesa in mesas)
        {
            var btn = CrearBotonMesa(mesa);
            _panelMesas.Controls.Add(btn);
        }

        _panelMesas.ResumeLayout();
    }

    private Button CrearBotonMesa(Mesa mesa)
    {
        var color = ColoresMesa.TryGetValue(mesa.Estado, out var c) ? c : Color.Gray;
        var label = LabelEstado.TryGetValue(mesa.Estado, out var l) ? l : mesa.Estado.ToString();

        var btn = new Button
        {
            Width     = 90,
            Height    = 80,
            Text      = $"{mesa.Resumido}\n{label}",
            FlatStyle = FlatStyle.Flat,
            BackColor = color,
            ForeColor = Color.White,
            Font      = new Font("Segoe UI", 7.5f, FontStyle.Bold),
            Cursor    = Cursors.Hand,
            Margin    = new Padding(4),
            Tag       = mesa
        };
        btn.FlatAppearance.BorderSize = 1;
        btn.FlatAppearance.BorderColor = Color.FromArgb(0, 0, 0, 60);
        btn.Click += async (_, _) => await SeleccionarMesaAsync(btn, mesa);
        return btn;
    }

    private async Task SeleccionarMesaAsync(Button btn, Mesa mesa)
    {
        if (_modo == ModoConsulta.Visual)
        {
            // BR-MESACONSULTA-001: permitir cambio de estado solo en no-ocupadas
            if (mesa.Estado == EstadoMesa.Ocupada)
            {
                MessageBox.Show("No se puede cambiar el estado de una mesa ocupada.",
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            await MostrarMenuEstadoAsync(btn, mesa);
        }
        else
        {
            // Modo Seleccionar / Mover: seleccionar la mesa destino
            MesaSeleccionada = mesa;
            // Deselect all, highlight selected
            foreach (Control ctrl in _panelMesas.Controls)
            {
                if (ctrl is Button b && b.Tag is Mesa m)
                    b.BackColor = ColoresMesa.TryGetValue(m.Estado, out var clr) ? clr : Color.Gray;
            }
            btn.BackColor           = Color.FromArgb(0, 80, 160);
            _btnAceptar.Enabled     = true;
        }
    }

    /// <summary>
    /// Muestra menú contextual para cambio de estado (modo Visual).
    /// BR-MESACONSULTA-002.
    /// </summary>
    private async Task MostrarMenuEstadoAsync(Button btn, Mesa mesaActual)
    {
        var ctx = new ContextMenuStrip();
        var estados = new[] { EstadoMesa.Libre, EstadoMesa.Reservada, EstadoMesa.Sucia, EstadoMesa.Bloqueada, EstadoMesa.FueraDeServicio };
        foreach (var est in estados)
        {
            if (est == mesaActual.Estado) continue;
            var lbl = LabelEstado.TryGetValue(est, out var l) ? l : est.ToString();
            var item = new ToolStripMenuItem($"Cambiar a → {lbl}");
            var estCapture = est;
            item.Click += async (_, _) =>
            {
                var r = await _cambiarEstadoHandler.HandleAsync(
                    new CambiarEstadoMesaCommand(mesaActual.CodigoMesa, estCapture));
                if (!r.EsExitoso)
                {
                    MessageBox.Show(r.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                // Actualizar visualmente el botón
                btn.BackColor = ColoresMesa.TryGetValue(estCapture, out var clr) ? clr : Color.Gray;
                var lbl2 = LabelEstado.TryGetValue(estCapture, out var l2) ? l2 : estCapture.ToString();
                btn.Text = $"{mesaActual.Resumido}\n{lbl2}";
                btn.Tag  = mesaActual;   // tag sigue referenciando la mesa, estado se refleja en el botón
            };
            ctx.Items.Add(item);
        }
        ctx.Show(btn, new Point(0, btn.Height));
        await Task.CompletedTask;
    }

    // ── Controles auxiliares ─────────────────────────────────────────────────

    private void CrearLeyendaEstados()
    {
        var flow = new FlowLayoutPanel
        {
            Dock          = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            Padding       = new Padding(6, 4, 6, 4)
        };
        var estados = new[] { EstadoMesa.Libre, EstadoMesa.Ocupada, EstadoMesa.Reservada, EstadoMesa.Sucia, EstadoMesa.Bloqueada, EstadoMesa.FueraDeServicio };
        foreach (var est in estados)
        {
            var color = ColoresMesa.TryGetValue(est, out var c) ? c : Color.Gray;
            var lbl   = LabelEstado.TryGetValue(est, out var l) ? l : est.ToString();
            var chip = new Label
            {
                Text      = $"  {lbl}  ",
                Height    = 22,
                AutoSize  = true,
                BackColor = color,
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 8),
                Padding   = new Padding(4, 2, 4, 2),
                Margin    = new Padding(3, 2, 3, 2),
                TextAlign = ContentAlignment.MiddleCenter
            };
            flow.Controls.Add(chip);
        }
        _panelEstados.Controls.Add(flow);
    }

    private void CrearContadores()
    {
        var flow = new FlowLayoutPanel
        {
            Dock          = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            Padding       = new Padding(8, 4, 8, 4)
        };
        flow.Controls.AddRange(new Control[]
        {
            new Label { Text = "Adultos:", AutoSize = true, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Segoe UI", 9), Margin = new Padding(0, 6, 4, 0) },
            _nudAdultos,
            new Label { Text = "Niños:", AutoSize = true, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Segoe UI", 9), Margin = new Padding(12, 6, 4, 0) },
            _nudNinios
        });
        _panelContadores.Controls.Add(flow);
    }

    private static Button CrearBotonAccion(string text, Color backColor) => new()
    {
        Text      = text,
        Width     = 100,
        Height    = 36,
        FlatStyle = FlatStyle.Flat,
        BackColor = backColor,
        ForeColor = Color.White,
        Font      = new Font("Segoe UI", 9, FontStyle.Bold),
        Cursor    = Cursors.Hand,
        Margin    = new Padding(4, 0, 0, 0)
    };
}
