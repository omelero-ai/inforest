using Inforest.Application.Interfaces;
using Inforest.Application.Pedidos;
using Inforest.Application.Turno;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Pedidos;

/// <summary>
/// Formulario de División de Pedidos.
/// Legacy: frmDivision.frm — División de Pedidos.
/// Reglas: BR-DIV-001, BR-DIV-002, BR-DIV-003, BR-DIV-004, BR-DIV-005.
/// </summary>
public sealed class FrmDivision : Form
{
    private readonly IniciarDivisionHandler _iniciarHandler;
    private readonly DisgregarItemHandler _disgregarHandler;
    private readonly CompartirPedidoHandler _compartirHandler;
    private readonly MoverItemHandler _moverHandler;
    private readonly ConfirmarDivisionHandler _confirmarHandler;
    private readonly IDiaContableService _diaContableService;
    private readonly ISessionService _sessionService;
    private readonly string _codigoPedido;
    private readonly string _codigoCaja;

    private SesionDivision? _sesion;
    private string? _pedidoDestinoSeleccionado;

    // ── controles ──
    private Label _lblInfoPedido = null!;
    private DataGridView _dgvPedidosDestino = null!;
    private DataGridView _dgvItemsOrigen = null!;
    private DataGridView _dgvItemsDestino = null!;
    private Label _lblMontoOrigen = null!;
    private Label _lblMontoDestino = null!;
    private Label _lblItemsOrigen = null!;

    /// <summary>True si la división fue confirmada exitosamente.</summary>
    public bool DivisionConfirmada { get; private set; }

    public FrmDivision(
        string codigoPedido,
        string codigoCaja,
        IniciarDivisionHandler iniciarHandler,
        DisgregarItemHandler disgregarHandler,
        CompartirPedidoHandler compartirHandler,
        MoverItemHandler moverHandler,
        ConfirmarDivisionHandler confirmarHandler,
        IDiaContableService diaContableService,
        ISessionService sessionService)
    {
        _codigoPedido = codigoPedido;
        _codigoCaja = codigoCaja;
        _iniciarHandler = iniciarHandler;
        _disgregarHandler = disgregarHandler;
        _compartirHandler = compartirHandler;
        _moverHandler = moverHandler;
        _confirmarHandler = confirmarHandler;
        _diaContableService = diaContableService;
        _sessionService = sessionService;

        Text = $"División de Pedidos — {codigoPedido}";
        Size = new Size(1200, 700);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        BuildLayout();
        Load += async (_, _) => await CargarSesionAsync();
    }

    // ── Layout ──────────────────────────────────────────────────────────────────

    private void BuildLayout()
    {
        // Top info bar
        _lblInfoPedido = new Label
        {
            Dock = DockStyle.Top,
            Height = 32,
            Padding = new Padding(8, 6, 0, 0),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(0, 84, 166),
            ForeColor = Color.White,
            Text = $"Pedido Origen: {_codigoPedido}"
        };

        // Center split — left panel (pedidos destino + items origen) | right panel (items destino)
        var splitMain = new SplitContainer
        {
            Dock = DockStyle.Fill,
            SplitterDistance = 580,
            IsSplitterFixed = false
        };

        // Left panel: top = pedidos destino, bottom = items origen
        var splitLeft = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = 180
        };

        // Pedidos destino (top left)
        var pnlPedidosDestino = BuildPanelConTitulo("Pedidos Destino", out _dgvPedidosDestino);
        _dgvPedidosDestino.SelectionChanged += DgvPedidosDestino_SelectionChanged;
        splitLeft.Panel1.Controls.Add(pnlPedidosDestino);

        // Items origen (bottom left)
        var pnlItemsOrigen = BuildPanelConTitulo("Items del Pedido Origen", out _dgvItemsOrigen);
        splitLeft.Panel2.Controls.Add(pnlItemsOrigen);

        splitMain.Panel1.Controls.Add(splitLeft);

        // Right panel: items destino
        var pnlItemsDestino = BuildPanelConTitulo("Items del Pedido Destino Seleccionado", out _dgvItemsDestino);
        splitMain.Panel2.Controls.Add(pnlItemsDestino);

        // Bottom button panel
        var pnlBotones = BuildPanelBotones();

        Controls.Add(splitMain);
        Controls.Add(pnlBotones);
        Controls.Add(_lblInfoPedido);
    }

    private static Panel BuildPanelConTitulo(string titulo, out DataGridView dgv)
    {
        var lbl = new Label
        {
            Dock = DockStyle.Top,
            Height = 24,
            Padding = new Padding(4, 4, 0, 0),
            BackColor = Color.FromArgb(52, 73, 94),
            ForeColor = Color.White,
            Text = titulo,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
        };

        dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            RowHeadersVisible = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 28
        };
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(41, 128, 185);
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);

        var panel = new Panel { Dock = DockStyle.Fill };
        panel.Controls.Add(dgv);
        panel.Controls.Add(lbl);
        return panel;
    }

    private Panel BuildPanelBotones()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 100,
            Padding = new Padding(8, 6, 8, 6),
            BackColor = Color.FromArgb(236, 240, 241)
        };

        // Monto + items info
        _lblItemsOrigen = new Label { Left = 10, Top = 8, Width = 200, Text = "Items origen: 0", AutoSize = false };
        _lblMontoOrigen = new Label { Left = 10, Top = 28, Width = 260, Text = "Total origen: 0.00", AutoSize = false, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        _lblMontoDestino = new Label { Left = 10, Top = 48, Width = 260, Text = "Total destino: 0.00", AutoSize = false };

        // Botones de items
        var btnDisgregar = CrearBoton("Disgregar", 280, 6, Color.FromArgb(231, 76, 60));
        var btnCompartir = CrearBoton("Compartir", 280, 34, Color.FromArgb(230, 126, 34));

        // Botones de movimiento
        var btnPasarUna = CrearBoton("Pasar ▶", 420, 6, Color.FromArgb(39, 174, 96));
        var btnQuitarUna = CrearBoton("◀ Quitar", 420, 34, Color.FromArgb(149, 165, 166));
        var btnPasarTodas = CrearBoton("Pasar todas ▶▶", 560, 6, Color.FromArgb(39, 174, 96));
        var btnQuitarTodas = CrearBoton("◀◀ Quitar todas", 560, 34, Color.FromArgb(149, 165, 166));

        // Botones de pedidos
        var btnAumentar = CrearBoton("+ Pedido", 720, 6, Color.FromArgb(52, 152, 219));
        var btnDisminuir = CrearBoton("- Pedido", 720, 34, Color.FromArgb(149, 165, 166));
        var btnCantidad = CrearBoton("Cantidad", 860, 6, Color.FromArgb(52, 152, 219));

        // Aceptar / Cancelar
        var btnAceptar = CrearBoton("✔ Aceptar", 1000, 6, Color.FromArgb(0, 84, 166));
        btnAceptar.Width = 150;
        var btnCancelar = CrearBoton("✖ Cancelar", 1000, 36, Color.FromArgb(192, 57, 43));
        btnCancelar.Width = 150;

        btnDisgregar.Click += async (_, _) => await DisgregarSeleccionadoAsync();
        btnCompartir.Click += async (_, _) => await CompartirAsync();
        btnPasarUna.Click += async (_, _) => await MoverItemAsync(toDestino: true, todas: false);
        btnQuitarUna.Click += async (_, _) => await MoverItemAsync(toDestino: false, todas: false);
        btnPasarTodas.Click += async (_, _) => await MoverItemAsync(toDestino: true, todas: true);
        btnQuitarTodas.Click += async (_, _) => await MoverItemAsync(toDestino: false, todas: true);
        btnAumentar.Click += (_, _) => AgregarPedidoDestino();
        btnDisminuir.Click += (_, _) => QuitarPedidoDestino();
        btnCantidad.Click += (_, _) => EstablecerCantidadPedidos();
        btnAceptar.Click += async (_, _) => await AceptarDivisionAsync();
        btnCancelar.Click += (_, _) => Close();

        panel.Controls.AddRange([
            _lblItemsOrigen, _lblMontoOrigen, _lblMontoDestino,
            btnDisgregar, btnCompartir,
            btnPasarUna, btnQuitarUna, btnPasarTodas, btnQuitarTodas,
            btnAumentar, btnDisminuir, btnCantidad,
            btnAceptar, btnCancelar
        ]);
        return panel;
    }

    private static Button CrearBoton(string texto, int left, int top, Color color)
        => new()
        {
            Text = texto,
            Left = left,
            Top = top,
            Width = 130,
            Height = 28,
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };

    // ── Carga inicial ────────────────────────────────────────────────────────────

    private async Task CargarSesionAsync()
    {
        var sesion = _sessionService.SesionActual;
        var correlativo = sesion?.CorrelativoAcceso.ToString() ?? "0";
        var obs = _codigoPedido;

        var result = await _iniciarHandler.HandleAsync(
            new IniciarDivisionCommand(_codigoPedido, _codigoCaja, correlativo, obs));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "División de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            Close();
            return;
        }

        _sesion = result.Valor;
        RefrescarTodo();
    }

    // ── Operaciones ──────────────────────────────────────────────────────────────

    private async Task DisgregarSeleccionadoAsync()
    {
        if (_sesion is null || _dgvItemsOrigen.SelectedRows.Count == 0) return;
        var tItem = _dgvItemsOrigen.SelectedRows[0].Cells["Item"].Value?.ToString();
        if (tItem is null) return;

        var item = _sesion.ItemsOrigen.FirstOrDefault(i => i.Item == tItem);
        if (item is null || item.Cantidad <= 1) { MessageBox.Show("El item no puede disgregarse (cantidad = 1)."); return; }

        // Pedir cantidad si lDisgrega (simplified: always ask)
        string? input = null;
        if (item.Cantidad > 2)
        {
            using var dlg = new Form { Text = "Cantidad a disgregar", Size = new Size(260, 130), FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent };
            var txt = new TextBox { Left = 10, Top = 10, Width = 220, Text = "1" };
            var btn = new Button { Text = "Aceptar", Left = 10, Top = 50, Width = 100, DialogResult = DialogResult.OK };
            dlg.Controls.AddRange([txt, btn]);
            dlg.AcceptButton = btn;
            if (dlg.ShowDialog(this) == DialogResult.OK) input = txt.Text;
        }

        double? cantidadPrimera = null;
        if (input is not null && double.TryParse(input, out double cant))
            cantidadPrimera = cant;

        var result = _disgregarHandler.HandleAsync(new DisgregarItemCommand(_sesion, tItem, cantidadPrimera));
        if (!result.EsExitoso) { MessageBox.Show(result.MensajeError); return; }
        RefrescarTodo();
    }

    private async Task CompartirAsync()
    {
        if (_sesion is null) return;
        var result = _compartirHandler.HandleAsync(_sesion);
        if (!result.EsExitoso) { MessageBox.Show(result.MensajeError); return; }
        RefrescarTodo();
        await Task.CompletedTask;
    }

    private async Task MoverItemAsync(bool toDestino, bool todas)
    {
        if (_sesion is null || _pedidoDestinoSeleccionado is null) return;

        if (todas)
        {
            if (toDestino) _sesion.MoverTodosADestino(_pedidoDestinoSeleccionado);
            else _sesion.RevertirTodosAOrigen(_pedidoDestinoSeleccionado);
        }
        else
        {
            var dgv = toDestino ? _dgvItemsOrigen : _dgvItemsDestino;
            if (dgv.SelectedRows.Count == 0) return;
            var tItem = dgv.SelectedRows[0].Cells["Item"].Value?.ToString();
            if (tItem is null) return;
            var result = _moverHandler.HandleAsync(new MoverItemCommand(_sesion, tItem, _pedidoDestinoSeleccionado, toDestino));
            if (!result.EsExitoso) { MessageBox.Show(result.MensajeError); return; }
        }
        RefrescarTodo();
        await Task.CompletedTask;
    }

    private void AgregarPedidoDestino()
    {
        if (_sesion is null) return;
        var sesion = _sessionService.SesionActual;
        string obs = _codigoPedido;
        _sesion.AgregarPedidoDestino(sesion?.CorrelativoAcceso.ToString() ?? "0", obs);
        RefrescarPedidosDestino();
    }

    private void QuitarPedidoDestino()
    {
        if (_sesion is null) return;
        bool ok = _sesion.QuitarUltimoPedidoDestino();
        if (!ok) { MessageBox.Show("El pedido destino tiene items asignados; no se puede eliminar."); return; }
        if (_pedidoDestinoSeleccionado is not null &&
            !_sesion.PedidosDestino.Any(p => p.CodigoTemporal == _pedidoDestinoSeleccionado))
            _pedidoDestinoSeleccionado = null;
        RefrescarTodo();
    }

    private void EstablecerCantidadPedidos()
    {
        if (_sesion is null) return;
        using var dlg = new Form { Text = "Cantidad de pedidos destino", Size = new Size(260, 130), FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent };
        var txt = new TextBox { Left = 10, Top = 10, Width = 220, Text = _sesion.PedidosDestino.Count.ToString() };
        var btn = new Button { Text = "Aceptar", Left = 10, Top = 50, Width = 100, DialogResult = DialogResult.OK };
        dlg.Controls.AddRange([txt, btn]);
        dlg.AcceptButton = btn;
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        if (!int.TryParse(txt.Text, out int cantidad) || cantidad < 1) return;

        var sesion = _sessionService.SesionActual;
        string obs = _codigoPedido;
        bool ok = _sesion.EstablecerCantidadPedidos(cantidad, sesion?.CorrelativoAcceso.ToString() ?? "0", obs);
        if (!ok) { MessageBox.Show("Algunos pedidos tienen items asignados; no se pueden eliminar."); return; }
        RefrescarTodo();
    }

    private async Task AceptarDivisionAsync()
    {
        if (_sesion is null) return;
        if (MessageBox.Show("¿Confirmar la división del pedido?", "División de Pedidos",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

        var diaContable = await _diaContableService.ObtenerDiaContableAsync(
            esAutomatico: false,
            horaCierre: "00:00",
            usuario: _sessionService.SesionActual?.CodigoUsuario ?? string.Empty);
        var fechaDC = diaContable.ToString("yyyyMMdd");

        var result = await _confirmarHandler.HandleAsync(
            new ConfirmarDivisionCommand(_sesion, fechaDC));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "División de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        DivisionConfirmada = true;
        Close();
    }

    // ── Refresco de UI ───────────────────────────────────────────────────────────

    private void RefrescarTodo()
    {
        RefrescarPedidosDestino();
        RefrescarItemsOrigen();
        RefrescarItemsDestino();
        ActualizarMontos();
    }

    private void RefrescarPedidosDestino()
    {
        if (_sesion is null) return;
        var dt = new System.Data.DataTable();
        dt.Columns.Add("CodigoTemporal");
        dt.Columns.Add("Observacion");
        dt.Columns.Add("Items", typeof(int));
        dt.Columns.Add("Total", typeof(decimal));

        foreach (var p in _sesion.PedidosDestino)
        {
            var items = _sesion.ItemsDestino(p.CodigoTemporal).ToList();
            dt.Rows.Add(p.CodigoTemporal, p.Observacion, items.Count, (decimal)items.Sum(i => i.Venta));
        }

        _dgvPedidosDestino.DataSource = dt;
        if (_dgvPedidosDestino.Columns.Count > 0)
        {
            _dgvPedidosDestino.Columns["CodigoTemporal"].HeaderText = "Pedido";
            _dgvPedidosDestino.Columns["Observacion"].HeaderText = "Observación";
            _dgvPedidosDestino.Columns["Items"].HeaderText = "Items";
            _dgvPedidosDestino.Columns["Total"].HeaderText = "Total";
        }
    }

    private void RefrescarItemsOrigen()
    {
        if (_sesion is null) return;
        BindItemsToGrid(_dgvItemsOrigen, _sesion.ItemsOrigen);
    }

    private void RefrescarItemsDestino()
    {
        if (_sesion is null || _pedidoDestinoSeleccionado is null) { _dgvItemsDestino.DataSource = null; return; }
        BindItemsToGrid(_dgvItemsDestino, _sesion.ItemsDestino(_pedidoDestinoSeleccionado));
    }

    private static void BindItemsToGrid(DataGridView dgv, IEnumerable<ItemDivision> items)
    {
        var dt = new System.Data.DataTable();
        dt.Columns.Add("Item");
        dt.Columns.Add("Producto");
        dt.Columns.Add("Precio", typeof(decimal));
        dt.Columns.Add("Cantidad", typeof(decimal));
        dt.Columns.Add("Total", typeof(decimal));

        foreach (var i in items)
            dt.Rows.Add(i.Item, i.NombreProducto, (decimal)i.PrecioVenta, (decimal)i.Cantidad, (decimal)i.Venta);

        dgv.DataSource = dt;
        if (dgv.Columns.Count > 0)
        {
            dgv.Columns["Item"].HeaderText = "Item";
            dgv.Columns["Producto"].HeaderText = "Producto";
            dgv.Columns["Precio"].HeaderText = "Precio";
            dgv.Columns["Cantidad"].HeaderText = "Cant.";
            dgv.Columns["Total"].HeaderText = "Total";
        }
    }

    private void ActualizarMontos()
    {
        if (_sesion is null) return;
        _lblItemsOrigen.Text = $"Items origen: {_sesion.ItemsOrigen.Count()}";
        _lblMontoOrigen.Text = $"Total origen: {_sesion.MontoOrigen:N2}";

        if (_pedidoDestinoSeleccionado is not null)
        {
            var totalDestino = (decimal)_sesion.ItemsDestino(_pedidoDestinoSeleccionado).Sum(i => i.Venta);
            _lblMontoDestino.Text = $"Total destino: {totalDestino:N2}";
        }
        else
        {
            _lblMontoDestino.Text = "Total destino: 0.00";
        }
    }

    private void DgvPedidosDestino_SelectionChanged(object? sender, EventArgs e)
    {
        if (_dgvPedidosDestino.SelectedRows.Count > 0)
        {
            _pedidoDestinoSeleccionado = _dgvPedidosDestino.SelectedRows[0]
                .Cells["CodigoTemporal"].Value?.ToString();
        }
        RefrescarItemsDestino();
        ActualizarMontos();
    }
}
