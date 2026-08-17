using Inforest.Application.Delivery;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario de búsqueda y selección de clientes delivery frecuentes.
/// <para>
/// Legacy: <c>frmBusquedaDelivery.frm</c> — buscador principal de TDELIVERY con
/// grilla filtrable (por teléfono, nombre o zona), panel de detalle con saldos
/// de fidelización, panel de estadísticas históricas ("Otros Datos"), y botones
/// para crear/editar cliente y seleccionar tienda/sucursal.
/// </para>
/// Legacy: <c>frmBusquedaDelivery.frm</c>, <c>TDELIVERY</c>, <c>vDelivery</c>,
/// <c>vZona</c>, <c>MPEDIDO</c>, <c>DPEDIDO</c>, <c>MDOCUMENTO</c>, <c>vTienda</c>.
/// Reglas: BR-DEL-036.
/// </summary>
public partial class FrmBusquedaDelivery : Form
{
    private readonly ObtenerClientesDeliveryBusquedaHandler _listarHandler;
    private readonly ObtenerDetalleClienteDeliveryHandler _detalleHandler;
    private readonly ObtenerEstadisticasClienteDeliveryHandler _statsHandler;
    private readonly ObtenerTiendasClienteDeliveryHandler _tiendasHandler;
    private readonly CrearClienteDeliveryHandler _crearHandler;
    private readonly ActualizarClienteDeliveryHandler _actualizarHandler;
    private readonly ActualizarFotoClienteDeliveryHandler _actualizarFotoHandler;
    private readonly BuscarClienteDeliveryHandler _buscarHandler;

    private IReadOnlyList<ClienteDeliveryBusquedaItem> _todosClientes = [];
    private IReadOnlyList<ClienteDeliveryBusquedaItem> _clientesFiltrados = [];

    // ── Resultado de selección ──────────────────────────────────────────────
    /// <summary>Código delivery del cliente seleccionado (null si se canceló).</summary>
    public string? CodigoDeliverySeleccionado { get; private set; }

    /// <summary>Nombre/razón del cliente seleccionado.</summary>
    public string? NombreSeleccionado { get; private set; }

    /// <summary>Teléfono del cliente seleccionado.</summary>
    public string? TelefonoSeleccionado { get; private set; }

    // ── Controles internos ──────────────────────────────────────────────────
    private DataGridView _grid = null!;
    private TextBox _txtBusqueda = null!;
    private Label _lblNombre = null!, _lblApellido = null!, _lblTelefono = null!;
    private Label _lblDireccion = null!, _lblZona = null!, _lblReferencia = null!;
    private Label _lblObservacion = null!, _lblDescuento = null!;
    private Label _lblAcumulado = null!, _lblUtilizado = null!, _lblDisponible = null!;
    private Label _lblTipoCliente = null!;
    private GroupBox _pnlOtrosDatos = null!;
    private Label _lblDesde = null!, _lblPedidos = null!, _lblTotal = null!;
    private Label _lblFechaUltima = null!, _lblUltimoDoc = null!, _lblMontoDoc = null!;
    private Label _lblRazonSocial = null!, _lblRuc = null!;
    private Button _btnOtrosDatos = null!;
    private bool _otrosDatosVisible;
    private int _diasHistorico = 30;

    public FrmBusquedaDelivery(
        ObtenerClientesDeliveryBusquedaHandler listarHandler,
        ObtenerDetalleClienteDeliveryHandler detalleHandler,
        ObtenerEstadisticasClienteDeliveryHandler statsHandler,
        ObtenerTiendasClienteDeliveryHandler tiendasHandler,
        CrearClienteDeliveryHandler crearHandler,
        ActualizarClienteDeliveryHandler actualizarHandler,
        ActualizarFotoClienteDeliveryHandler actualizarFotoHandler,
        BuscarClienteDeliveryHandler buscarHandler)
    {
        _listarHandler = listarHandler;
        _detalleHandler = detalleHandler;
        _statsHandler = statsHandler;
        _tiendasHandler = tiendasHandler;
        _crearHandler = crearHandler;
        _actualizarHandler = actualizarHandler;
        _actualizarFotoHandler = actualizarFotoHandler;
        _buscarHandler = buscarHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Búsqueda de Cliente Delivery";
        Size = new Size(1020, 680);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        KeyPreview = true;

        // ── Barra de búsqueda ───────────────────────────────────────────────
        var pnlBusqueda = new Panel { Dock = DockStyle.Top, Height = 42, Padding = new Padding(8, 8, 8, 4) };
        var lblBuscar = new Label { Text = "Buscar:", Left = 8, Top = 11, Width = 55, TextAlign = ContentAlignment.MiddleLeft };
        _txtBusqueda = new TextBox { Left = 68, Top = 8, Width = 300, Name = "txtBusqueda" };
        var lblCriterio = new Label { Text = "Columna:", Left = 378, Top = 11, Width = 65, TextAlign = ContentAlignment.MiddleLeft };
        var cboCriterio = new ComboBox
        {
            Name = "cboCriterio",
            Left = 448, Top = 8, Width = 130,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        cboCriterio.Items.AddRange(["Teléfono", "Cliente", "Zona"]);
        cboCriterio.SelectedIndex = 0;
        pnlBusqueda.Controls.AddRange([lblBuscar, _txtBusqueda, lblCriterio, cboCriterio]);

        // ── Grid principal ──────────────────────────────────────────────────
        _grid = new DataGridView
        {
            Name = "grdClientes",
            Dock = DockStyle.None,
            Location = new Point(8, 50),
            Size = new Size(580, 500),
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            MultiSelect = false,
            RowHeadersVisible = false,
        };
        _grid.Columns.AddRange(
            new DataGridViewTextBoxColumn { Name = "colCodigo", HeaderText = "Código", Width = 80, DataPropertyName = "CodigoDelivery" },
            new DataGridViewTextBoxColumn { Name = "colTelefono", HeaderText = "Teléfono", Width = 120, DataPropertyName = "Telefono" },
            new DataGridViewTextBoxColumn { Name = "colCliente", HeaderText = "Cliente", Width = 230, DataPropertyName = "Cliente" },
            new DataGridViewTextBoxColumn { Name = "colZona", HeaderText = "Zona", Width = 140, DataPropertyName = "Zona" }
        );

        // ── Panel de detalle (derecha) ──────────────────────────────────────
        var pnlDetalle = new GroupBox
        {
            Text = "Datos del Cliente",
            Location = new Point(598, 50),
            Size = new Size(400, 500),
        };

        int dy = 18;
        int labelWidth = 95;
        int valueLeft = 100;

        Label MkLbl(string text)
        {
            var l = new Label { Text = text, Left = 6, Top = dy, Width = labelWidth, TextAlign = ContentAlignment.MiddleRight };
            return l;
        }
        Label MkVal(ref Label field)
        {
            field = new Label { Left = valueLeft, Top = dy, Width = 290, AutoSize = false, BorderStyle = BorderStyle.FixedSingle, BackColor = SystemColors.Window };
            dy += 28;
            return field;
        }

        var lblNombreHdr = MkLbl("Nombre:");       var valNombre = MkVal(ref _lblNombre);
        var lblApellidoHdr = MkLbl("Apellido:");   var valApellido = MkVal(ref _lblApellido);
        var lblTelefonoHdr = MkLbl("Teléfono:");   var valTelefono = MkVal(ref _lblTelefono);
        var lblDirecHdr = MkLbl("Dirección:");     var valDireccion = MkVal(ref _lblDireccion);
        var lblZonaHdr = MkLbl("Zona:");           var valZona = MkVal(ref _lblZona);
        var lblRefHdr = MkLbl("Referencia:");      var valReferencia = MkVal(ref _lblReferencia);
        var lblObsHdr = MkLbl("Observación:");     var valObservacion = MkVal(ref _lblObservacion);
        var lblDescuentoHdr = MkLbl("Descuento:"); var valDescuento = MkVal(ref _lblDescuento);
        var lblAcumHdr = MkLbl("Acumulado:");      var valAcumulado = MkVal(ref _lblAcumulado);
        var lblUtilHdr = MkLbl("Utilizado:");      var valUtilizado = MkVal(ref _lblUtilizado);
        var lblDispHdr = MkLbl("Disponible:");     var valDisponible = MkVal(ref _lblDisponible);
        var lblTipoHdr = MkLbl("Tipo Cliente:");   var valTipoCliente = MkVal(ref _lblTipoCliente);

        pnlDetalle.Controls.AddRange([
            lblNombreHdr, valNombre, lblApellidoHdr, valApellido,
            lblTelefonoHdr, valTelefono, lblDirecHdr, valDireccion,
            lblZonaHdr, valZona, lblRefHdr, valReferencia,
            lblObsHdr, valObservacion, lblDescuentoHdr, valDescuento,
            lblAcumHdr, valAcumulado, lblUtilHdr, valUtilizado,
            lblDispHdr, valDisponible, lblTipoHdr, valTipoCliente
        ]);

        // ── Panel "Otros Datos" (oculto por defecto) ────────────────────────
        _pnlOtrosDatos = new GroupBox
        {
            Text = "Estadísticas (últimos N días)",
            Location = new Point(598, 50),
            Size = new Size(400, 500),
            Visible = false
        };

        int od = 18;
        Label OdLbl(string t) => new() { Text = t, Left = 6, Top = od, Width = 105, TextAlign = ContentAlignment.MiddleRight };
        Label OdVal(ref Label f) { f = new Label { Left = 112, Top = od, Width = 278, AutoSize = false, BorderStyle = BorderStyle.FixedSingle, BackColor = SystemColors.Window }; od += 28; return f; }

        var odDesdeHdr = OdLbl("Desde:");          var odDesde = OdVal(ref _lblDesde);
        var odPedidosHdr = OdLbl("Pedidos:");       var odPedidos = OdVal(ref _lblPedidos);
        var odTotalHdr = OdLbl("Total:");           var odTotal = OdVal(ref _lblTotal);
        var odFechaHdr = OdLbl("Última Compra:");   var odFecha = OdVal(ref _lblFechaUltima);
        var odDocHdr = OdLbl("Último Doc:");        var odDoc = OdVal(ref _lblUltimoDoc);
        var odMontoHdr = OdLbl("Monto:");           var odMonto = OdVal(ref _lblMontoDoc);
        var odRsHdr = OdLbl("Razón Social:");       var odRs = OdVal(ref _lblRazonSocial);
        var odRucHdr = OdLbl("RUC:");               var odRuc = OdVal(ref _lblRuc);

        _pnlOtrosDatos.Controls.AddRange([
            odDesdeHdr, odDesde, odPedidosHdr, odPedidos, odTotalHdr, odTotal,
            odFechaHdr, odFecha, odDocHdr, odDoc, odMontoHdr, odMonto,
            odRsHdr, odRs, odRucHdr, odRuc
        ]);

        // ── Barra de botones inferior ───────────────────────────────────────
        var pnlBotones = new Panel { Dock = DockStyle.Bottom, Height = 48 };
        int bx = 8;
        Button Btn(string text, int width = 110)
        {
            var b = new Button { Text = text, Left = bx, Top = 10, Width = width, Height = 30 };
            bx += width + 6;
            pnlBotones.Controls.Add(b);
            return b;
        }

        var btnNuevo = Btn("Nuevo Cliente");
        var btnModificar = Btn("Modificar");
        _btnOtrosDatos = Btn("Otros Datos");
        var btnTiendas = Btn("Tiendas", 90);
        var btnSeleccionar = Btn("Seleccionar", 100);
        var btnCancelar = Btn("Cancelar", 90);

        // ── Componer formulario ─────────────────────────────────────────────
        Controls.AddRange([pnlBusqueda, _grid, pnlDetalle, _pnlOtrosDatos, pnlBotones]);

        // ── Eventos ─────────────────────────────────────────────────────────
        _txtBusqueda.TextChanged += (_, _) => FiltrarGrilla(cboCriterio.SelectedIndex);
        cboCriterio.SelectedIndexChanged += (_, _) => FiltrarGrilla(cboCriterio.SelectedIndex);

        _grid.SelectionChanged += async (_, _) => await CargarDetalleAsync();
        _grid.CellDoubleClick += (_, _) => ConfirmarSeleccion();

        btnNuevo.Click += async (_, _) => await AbrirNuevoClienteAsync(esNuevo: true);
        btnModificar.Click += async (_, _) => await AbrirNuevoClienteAsync(esNuevo: false);
        _btnOtrosDatos.Click += async (_, _) => await ToggleOtrosDatosAsync();
        btnTiendas.Click += async (_, _) => await MostrarTiendasAsync();
        btnSeleccionar.Click += (_, _) => ConfirmarSeleccion();
        btnCancelar.Click += (_, _) => CerrarCancelado();

        Load += async (_, _) => await CargarClientesAsync();
        KeyDown += OnKeyDown;
    }

    // ── Carga inicial ────────────────────────────────────────────────────────

    private async Task CargarClientesAsync()
    {
        var result = await _listarHandler.HandleAsync(new ObtenerClientesDeliveryBusquedaQuery());
        if (!result.EsExitoso) return;
        _todosClientes = result.Valor!;
        _clientesFiltrados = _todosClientes;
        BindGrid();
    }

    private void BindGrid()
    {
        var bs = new BindingSource { DataSource = _clientesFiltrados.ToList() };
        _grid.DataSource = bs;
    }

    private void FiltrarGrilla(int criterio)
    {
        var texto = _txtBusqueda.Text.Trim().ToUpperInvariant();
        if (string.IsNullOrEmpty(texto))
        {
            _clientesFiltrados = _todosClientes;
        }
        else
        {
            _clientesFiltrados = criterio switch
            {
                0 => _todosClientes.Where(c => (c.Telefono ?? "").Contains(texto, StringComparison.OrdinalIgnoreCase)).ToList(),
                1 => _todosClientes.Where(c => c.Cliente.Contains(texto, StringComparison.OrdinalIgnoreCase)).ToList(),
                2 => _todosClientes.Where(c => (c.Zona ?? "").Contains(texto, StringComparison.OrdinalIgnoreCase)).ToList(),
                _ => _todosClientes
            };
        }
        BindGrid();
    }

    // ── Detalle del cliente seleccionado ────────────────────────────────────

    private async Task CargarDetalleAsync()
    {
        var codigo = ObtenerCodigoSeleccionado();
        if (string.IsNullOrEmpty(codigo)) return;

        var result = await _detalleHandler.HandleAsync(new ObtenerDetalleClienteDeliveryQuery(codigo));
        if (!result.EsExitoso || result.Valor is null) return;

        var d = result.Valor;
        _lblNombre.Text     = d.Nombre ?? "";
        _lblApellido.Text   = d.Apellido ?? "";
        _lblTelefono.Text   = d.Telefono ?? "";
        _lblDireccion.Text  = d.Direccion ?? "";
        _lblZona.Text       = d.Zona ?? "";
        _lblReferencia.Text = d.Referencia ?? "";
        _lblObservacion.Text = d.Observacion ?? "";
        _lblDescuento.Text  = d.Descuento.ToString("N2");
        _lblAcumulado.Text  = d.Acumulado.ToString("N2");
        _lblUtilizado.Text  = d.Utilizado.ToString("N2");
        _lblDisponible.Text = d.Disponible.ToString("N2");
        _lblTipoCliente.Text = d.TipoCliente ?? "";

        // Si "Otros Datos" visible, también recargamos las estadísticas
        if (_otrosDatosVisible)
            await CargarEstadisticasAsync(codigo);
    }

    private void LimpiarDetalle()
    {
        foreach (var lbl in new[] { _lblNombre, _lblApellido, _lblTelefono, _lblDireccion,
                     _lblZona, _lblReferencia, _lblObservacion, _lblDescuento,
                     _lblAcumulado, _lblUtilizado, _lblDisponible, _lblTipoCliente })
            lbl.Text = "";
    }

    // ── "Otros Datos" ───────────────────────────────────────────────────────

    private async Task ToggleOtrosDatosAsync()
    {
        _otrosDatosVisible = !_otrosDatosVisible;
        var pnlDetalle = Controls.OfType<GroupBox>().FirstOrDefault(g => g.Text == "Datos del Cliente");
        if (pnlDetalle != null) pnlDetalle.Visible = !_otrosDatosVisible;
        _pnlOtrosDatos.Visible = _otrosDatosVisible;
        _btnOtrosDatos.Text = _otrosDatosVisible ? "Datos Principales" : "Otros Datos";

        if (_otrosDatosVisible)
        {
            var codigo = ObtenerCodigoSeleccionado();
            if (!string.IsNullOrEmpty(codigo))
                await CargarEstadisticasAsync(codigo);
        }
    }

    private async Task CargarEstadisticasAsync(string codigoDelivery)
    {
        var result = await _statsHandler.HandleAsync(
            new ObtenerEstadisticasClienteDeliveryQuery(codigoDelivery, _diasHistorico));
        if (!result.EsExitoso) return;

        var s = result.Valor;
        _lblDesde.Text        = s.Desde?.ToString("dd/MM/yyyy") ?? "";
        _lblPedidos.Text      = s.NumeroPedidos.ToString("N0");
        _lblTotal.Text        = s.TotalVenta.ToString("N2");
        _lblFechaUltima.Text  = s.FechaUltimaCompra?.ToString("dd/MM/yyyy") ?? "";
        _lblUltimoDoc.Text    = s.UltimoDocumento ?? "";
        _lblMontoDoc.Text     = s.MontoUltimoDoc.ToString("N2");
        _lblRazonSocial.Text  = s.RazonSocial ?? "";
        _lblRuc.Text          = s.Ruc ?? "";
    }

    // ── Tiendas ──────────────────────────────────────────────────────────────

    private async Task MostrarTiendasAsync()
    {
        var codigo = ObtenerCodigoSeleccionado();
        if (string.IsNullOrEmpty(codigo)) return;

        var result = await _tiendasHandler.HandleAsync(new ObtenerTiendasClienteDeliveryQuery(codigo));
        if (!result.EsExitoso || !result.Valor!.Any())
        {
            MessageBox.Show("Este cliente no tiene tiendas/sucursales registradas.", "Tiendas", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dlg = new FrmBusquedaTiendasDelivery(result.Valor!);
        if (dlg.ShowDialog(this) == DialogResult.OK && dlg.TiendaSeleccionada is not null)
        {
            // Información de tienda disponible en dlg.TiendaSeleccionada
        }
    }

    // ── Crear / Modificar cliente ────────────────────────────────────────────

    private async Task AbrirNuevoClienteAsync(bool esNuevo)
    {
        using var dlg = new NuevoDeliveryForm(_crearHandler, _actualizarHandler, _actualizarFotoHandler, _buscarHandler)
        {
            Text = esNuevo ? "Agregar Cliente Frecuente" : "Modificar Cliente Delivery"
        };
        dlg.ShowDialog(this);

        // Recargar la grilla con los datos actualizados
        await CargarClientesAsync();
    }

    // ── Selección y cierre ───────────────────────────────────────────────────

    private void ConfirmarSeleccion()
    {
        var codigo = ObtenerCodigoSeleccionado();
        if (string.IsNullOrEmpty(codigo)) return;

        var row = _clientesFiltrados.FirstOrDefault(c => c.CodigoDelivery == codigo);
        CodigoDeliverySeleccionado = codigo;
        NombreSeleccionado         = row?.Cliente;
        TelefonoSeleccionado       = row?.Telefono;
        DialogResult = DialogResult.OK;
        Close();
    }

    private void CerrarCancelado()
    {
        CodigoDeliverySeleccionado = null;
        DialogResult = DialogResult.Cancel;
        Close();
    }

    private string? ObtenerCodigoSeleccionado()
    {
        if (_grid.SelectedRows.Count == 0) return null;
        return _grid.SelectedRows[0].Cells["colCodigo"].Value?.ToString();
    }

    private void OnKeyDown(object? sender, KeyEventArgs e)
    {
        switch (e.KeyCode)
        {
            case Keys.Enter:
                ConfirmarSeleccion();
                e.Handled = true;
                break;
            case Keys.Escape:
                CerrarCancelado();
                e.Handled = true;
                break;
        }
    }
}
