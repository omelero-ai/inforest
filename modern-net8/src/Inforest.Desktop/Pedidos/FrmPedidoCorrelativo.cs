using Inforest.Application.Pedidos;

namespace Inforest.Desktop.Pedidos;

/// <summary>
/// Correlativo de pedidos — consulta por rango de fechas con detalle de ítems y documentos.
/// Legacy: <c>frmPedidoCorrelativo.frm</c> + <c>frmPedidoDetalle.frm</c> + <c>frmPedidoAnterior.frm</c>.
/// Regla BR-CORRPEDIDO-001: rango de fechas, vPedidoCorrelativo, detalle vPedidoDetalle, docs vDocumentoAgrupado.
/// SQL: vPedidoCorrelativo, vPedidoDetalle, vDocumentoAgrupado.
/// </summary>
public sealed class FrmPedidoCorrelativo : Form
{
    // ── Design tokens ─────────────────────────────────────────────────────────
    private static readonly Color Surface     = Color.FromArgb(245, 246, 247);
    private static readonly Color Header      = Color.FromArgb(40, 60, 90);
    private static readonly Color Accent      = Color.FromArgb(0, 122, 204);
    private static readonly Color RowAlt      = Color.FromArgb(240, 244, 250);
    private static readonly Color Panel       = Color.FromArgb(235, 236, 240);
    private static readonly Font  LabelFont   = new("Segoe UI", 9f);
    private static readonly Font  HeaderFont  = new("Segoe UI", 10.5f, FontStyle.Bold);
    private static readonly Font  GridFont    = new("Segoe UI", 8.5f);

    // ── Dependencies ──────────────────────────────────────────────────────────
    private readonly ObtenerPedidosCorrelativoHandler         _correlativoHandler;
    private readonly ObtenerDetallePedidoExtendidoHandler     _detalleHandler;
    private readonly ObtenerDocumentosAgrupadosPedidoHandler  _documentosHandler;

    // ── State ─────────────────────────────────────────────────────────────────
    private IReadOnlyList<PedidoCorrelativoVista>  _pedidos   = [];
    private string?                                _pedidoSel;

    // ── Top filter bar ────────────────────────────────────────────────────────
    private readonly DateTimePicker _dtpDesde = new() { Width = 110, Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddDays(-1) };
    private readonly DateTimePicker _dtpHasta = new() { Width = 110, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
    private readonly Button _btnBuscar   = MakeBtn("Buscar",    Accent,                         Color.White);
    private readonly Button _btnDetalle  = MakeBtn("Ítems",     Color.FromArgb(70, 130, 180),   Color.White);
    private readonly Button _btnDocs     = MakeBtn("Docs",      Color.FromArgb(100, 140, 60),   Color.White);
    private readonly Button _btnSalir    = MakeBtn("Salir",     Color.FromArgb(108, 117, 125),  Color.White);
    private readonly Label  _lblTotal    = new() { AutoSize = true, Font = LabelFont, ForeColor = Color.FromArgb(60, 80, 120) };

    // ── Main grid (pedidos) ───────────────────────────────────────────────────
    private readonly DataGridView _dgvPedidos = new()
    {
        Dock              = DockStyle.Fill,
        ReadOnly          = true,
        AutoGenerateColumns = false,
        SelectionMode     = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false,
        RowHeadersVisible = false,
        BorderStyle       = BorderStyle.None,
        Font              = GridFont,
        BackgroundColor   = Surface,
        GridColor         = Color.FromArgb(220, 222, 226),
    };

    // ── Detail tabs ───────────────────────────────────────────────────────────
    private readonly DataGridView _dgvItems = new()
    {
        Dock              = DockStyle.Fill,
        ReadOnly          = true,
        AutoGenerateColumns = false,
        SelectionMode     = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false,
        RowHeadersVisible = false,
        BorderStyle       = BorderStyle.None,
        Font              = GridFont,
        BackgroundColor   = Surface,
        GridColor         = Color.FromArgb(220, 222, 226),
    };

    private readonly DataGridView _dgvDocumentos = new()
    {
        Dock              = DockStyle.Fill,
        ReadOnly          = true,
        AutoGenerateColumns = false,
        SelectionMode     = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false,
        RowHeadersVisible = false,
        BorderStyle       = BorderStyle.None,
        Font              = GridFont,
        BackgroundColor   = Surface,
        GridColor         = Color.FromArgb(220, 222, 226),
    };

    private readonly TabControl _tabs = new() { Dock = DockStyle.Fill };
    private readonly Label _lblPedidoSel = new() { AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };

    public FrmPedidoCorrelativo(
        ObtenerPedidosCorrelativoHandler correlativoHandler,
        ObtenerDetallePedidoExtendidoHandler detalleHandler,
        ObtenerDocumentosAgrupadosPedidoHandler documentosHandler)
    {
        _correlativoHandler = correlativoHandler;
        _detalleHandler     = detalleHandler;
        _documentosHandler  = documentosHandler;

        Text            = "Correlativo de Pedidos";
        Width           = 1200;
        Height          = 780;
        StartPosition   = FormStartPosition.CenterScreen;
        BackColor       = Surface;
        MinimumSize     = new Size(900, 600);
        FormBorderStyle = FormBorderStyle.Sizable;

        BuildColumns();
        BuildLayout();
        WireEvents();
    }

    // ── Column definitions ────────────────────────────────────────────────────

    private void BuildColumns()
    {
        _dgvPedidos.Columns.AddRange(
            Col("CodigoPedido",  "Pedido",     90),
            Col("Caja",          "Caja",       60),
            Col("Turno",         "Turno",      80),
            ColDate("Fecha",     "Fecha",      90),
            Col("Mesa",          "Mesa",       80),
            Col("Mozo",          "Mozo",      120),
            ColNum("Venta",      "Total",      90, "N2"),
            Col("TipoPedido",    "Tipo",       80),
            Col("Estado",        "Estado",     90),
            Col("Documento",     "Documento", 140),
            Col("Observacion",   "Obs.",      130));

        _dgvItems.Columns.AddRange(
            Col("Item",          "Ítem",       55),
            Col("Producto",      "Producto",  240),
            ColNum("PrecioVenta", "P.Venta",   90, "N2"),
            ColNum("Cantidad",   "Cant.",      70, "N2"),
            ColNum("SubTotal",   "SubTotal",   90, "N2"),
            Col("Documento",     "Documento", 140));

        _dgvDocumentos.Columns.AddRange(
            Col("Documento",        "Documento",     140),
            ColNum("Venta",         "Monto",          90, "N2"),
            Col("Estado",           "Estado",          90),
            Col("Turno",            "Turno",           80),
            Col("ObservacionAnulacion", "Obs. Anul.", 160),
            Col("UsuarioAnulado",   "Usuario Anul.",   110),
            ColDate("FechaAnulacion","Fec. Anul.",      90));
    }

    // ── Layout ────────────────────────────────────────────────────────────────

    private void BuildLayout()
    {
        // ── Header bar ───
        var header = new Panel
        {
            Dock      = DockStyle.Top,
            Height    = 42,
            BackColor = Header,
            Padding   = new Padding(10, 0, 10, 0),
        };
        var lblTitle = new Label
        {
            Text      = "Correlativo de Pedidos",
            Font      = HeaderFont,
            ForeColor = Color.White,
            Dock      = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        header.Controls.Add(lblTitle);

        // ── Filter bar ───
        var filterPanel = new Panel { Dock = DockStyle.Top, Height = 42, BackColor = Panel, Padding = new Padding(8, 6, 8, 4) };
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
        flow.Controls.AddRange(new Control[]
        {
            MakeLbl("Desde:"), _dtpDesde,
            MakeSpacer(8),
            MakeLbl("Hasta:"), _dtpHasta,
            MakeSpacer(8),
            _btnBuscar, MakeSpacer(16),
            _btnDetalle, _btnDocs,
            MakeSpacer(16),
            _lblTotal,
        });
        filterPanel.Controls.Add(flow);

        // ── Main splitter ───
        var split = new SplitContainer
        {
            Dock        = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = 340,
            Panel1MinSize    = 150,
            Panel2MinSize    = 100,
        };

        // Top: orders grid
        _dgvPedidos.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = RowAlt };
        split.Panel1.Controls.Add(_dgvPedidos);

        // Bottom: tabs (items + docs)
        var tabItems = new TabPage("Ítems del Pedido") { BackColor = Surface };
        tabItems.Controls.Add(_dgvItems);
        var tabDocs = new TabPage("Documentos") { BackColor = Surface };
        tabDocs.Controls.Add(_dgvDocumentos);
        _tabs.TabPages.AddRange([tabItems, tabDocs]);

        var bottomPanel = new Panel { Dock = DockStyle.Fill };
        var lblSelPanel = new Panel { Dock = DockStyle.Top, Height = 24, BackColor = Panel, Padding = new Padding(8, 4, 0, 0) };
        _lblPedidoSel.Text = "Pedido seleccionado: —";
        lblSelPanel.Controls.Add(_lblPedidoSel);
        bottomPanel.Controls.Add(_tabs);
        bottomPanel.Controls.Add(lblSelPanel);
        split.Panel2.Controls.Add(bottomPanel);

        // ── Footer bar ───
        var footer = new Panel { Dock = DockStyle.Bottom, Height = 44, BackColor = Panel, Padding = new Padding(8, 6, 8, 4) };
        var footFlow = new FlowLayoutPanel { Dock = DockStyle.Right, WrapContents = false };
        footFlow.Controls.Add(_btnSalir);
        footer.Controls.Add(footFlow);

        Controls.AddRange([header, filterPanel, split, footer]);
    }

    // ── Events ───────────────────────────────────────────────────────────────

    private void WireEvents()
    {
        _btnBuscar.Click  += async (_, _) => await CargarPedidosAsync();
        _btnDetalle.Click += async (_, _) => await CargarDetalleAsync();
        _btnDocs.Click    += async (_, _) => await CargarDocumentosAsync();
        _btnSalir.Click   += (_, _)        => Close();

        _dgvPedidos.SelectionChanged += async (_, _) => await OnPedidoSeleccionadoAsync();
        _dgvPedidos.CellDoubleClick  += async (_, _) => await CargarDetalleAsync();
        Load += async (_, _) => await CargarPedidosAsync();
    }

    // ── Operations ───────────────────────────────────────────────────────────

    private async Task CargarPedidosAsync()
    {
        _dgvPedidos.DataSource = null;
        _lblTotal.Text = "Buscando…";
        try
        {
            var query  = new ObtenerPedidosCorrelativoQuery(
                DateOnly.FromDateTime(_dtpDesde.Value),
                DateOnly.FromDateTime(_dtpHasta.Value));
            var result = await _correlativoHandler.HandleAsync(query);
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _pedidos = result.Valor!;
            _dgvPedidos.DataSource = _pedidos.ToList();
            _lblTotal.Text = $"{_pedidos.Count} pedido(s)";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async Task OnPedidoSeleccionadoAsync()
    {
        if (_dgvPedidos.CurrentRow?.DataBoundItem is not PedidoCorrelativoVista p) return;
        _pedidoSel = p.CodigoPedido;
        _lblPedidoSel.Text = $"Pedido seleccionado: {_pedidoSel}  |  Mesa: {p.Mesa}  |  Total: {p.Venta:N2}";

        // Auto-load active tab
        if (_tabs.SelectedIndex == 0)
            await CargarDetalleAsync();
        else
            await CargarDocumentosAsync();
    }

    /// <summary>
    /// Carga los ítems del pedido seleccionado (tab Ítems).
    /// Legacy: frmPedidoDetalle.frm — vPedidoDetalle WHERE tCodigoPedido = @ped.
    /// </summary>
    private async Task CargarDetalleAsync()
    {
        if (_pedidoSel is null) return;
        _tabs.SelectedIndex = 0;
        _dgvItems.DataSource = null;
        try
        {
            var result = await _detalleHandler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery(_pedidoSel));
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _dgvItems.DataSource = result.Valor!.Items.ToList();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// Carga los documentos agrupados del pedido seleccionado (tab Documentos).
    /// Legacy: frmPedidoAnterior.frm — vDocumentoAgrupado WHERE tCodigoPedido = @ped.
    /// </summary>
    private async Task CargarDocumentosAsync()
    {
        if (_pedidoSel is null) return;
        _tabs.SelectedIndex = 1;
        _dgvDocumentos.DataSource = null;
        try
        {
            var result = await _documentosHandler.HandleAsync(new ObtenerDocumentosAgrupadosPedidoQuery(_pedidoSel));
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _dgvDocumentos.DataSource = result.Valor!.ToList();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ── Factory helpers ───────────────────────────────────────────────────────

    private static Button MakeBtn(string text, Color back, Color fore) => new()
    {
        Text      = text,
        Width     = 80,
        Height    = 28,
        BackColor = back,
        ForeColor = fore,
        FlatStyle = FlatStyle.Flat,
        Font      = new Font("Segoe UI", 8.5f),
        Margin    = new Padding(4, 0, 0, 0),
        FlatAppearance = { BorderSize = 0 },
    };

    private static Label MakeLbl(string text) => new()
    {
        Text      = text,
        AutoSize  = true,
        Font      = LabelFont,
        Margin    = new Padding(0, 6, 2, 0),
    };

    private static Panel MakeSpacer(int w) => new() { Width = w, Height = 1 };

    private static DataGridViewTextBoxColumn Col(string name, string header, int width) => new()
    {
        DataPropertyName = name,
        HeaderText       = header,
        Width            = width,
        SortMode         = DataGridViewColumnSortMode.Automatic,
    };

    private static DataGridViewTextBoxColumn ColNum(string name, string header, int width, string fmt) => new()
    {
        DataPropertyName = name,
        HeaderText       = header,
        Width            = width,
        DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = fmt },
        SortMode         = DataGridViewColumnSortMode.Automatic,
    };

    private static DataGridViewTextBoxColumn ColDate(string name, string header, int width) => new()
    {
        DataPropertyName = name,
        HeaderText       = header,
        Width            = width,
        DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy" },
        SortMode         = DataGridViewColumnSortMode.Automatic,
    };
}
