using Inforest.Application.Delivery;
using Inforest.Application.Pedidos;
using Inforest.Domain.Entities.Delivery;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Seguimiento de pedidos delivery entregados.
/// Legacy: <c>frmPedidoDeliveryNo.frm</c>, <c>vDespachador</c>, <c>vPedidoDetalle</c>.
/// Regla: BR-DEL-038.
/// </summary>
public sealed class FrmPedidoDeliveryNo : Form
{
    private readonly ObtenerPedidosSeguimientoDeliveryEntregadosHandler _seguimientoHandler;
    private readonly ObtenerDetallePedidoExtendidoHandler _detalleHandler;

    private readonly BindingSource _binding = new();
    private IReadOnlyList<PedidoDeliverySeguimiento> _todos = [];
    private List<PedidoDeliverySeguimiento> _filtrados = [];

    private DataGridView _gridCabecera = null!;
    private DataGridView _gridDetalle = null!;
    private ComboBox _cboCriterio = null!;
    private TextBox _txtFiltro = null!;
    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private Label _lblRegistro = null!;
    private Label _lblReferencia = null!;
    private Label _lblDireccion = null!;
    private Label _lblZona = null!;

    public FrmPedidoDeliveryNo(
        ObtenerPedidosSeguimientoDeliveryEntregadosHandler seguimientoHandler,
        ObtenerDetallePedidoExtendidoHandler detalleHandler)
    {
        _seguimientoHandler = seguimientoHandler;
        _detalleHandler = detalleHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Seguimientos de Pedidos Delivery Entregados";
        WindowState = FormWindowState.Maximized;
        MinimumSize = new Size(1200, 680);
        BackColor = Color.FromArgb(245, 246, 247);

        var panelTop = new Panel
        {
            Dock = DockStyle.Top,
            Height = 88,
            Padding = new Padding(12, 10, 12, 8),
            BackColor = Color.White
        };

        panelTop.Controls.Add(new Label { Text = "Fecha inicial", Left = 12, Top = 14, Width = 88, Font = new Font("Segoe UI", 9, FontStyle.Bold) });
        _dtpInicio = new DateTimePicker { Left = 104, Top = 10, Width = 130, Value = DateTime.Today };
        panelTop.Controls.Add(_dtpInicio);

        panelTop.Controls.Add(new Label { Text = "Fecha final", Left = 245, Top = 14, Width = 80, Font = new Font("Segoe UI", 9, FontStyle.Bold) });
        _dtpFin = new DateTimePicker { Left = 326, Top = 10, Width = 130, Value = DateTime.Today };
        panelTop.Controls.Add(_dtpFin);

        var btnProcesar = CrearBoton("Procesar", 470, Color.FromArgb(14, 116, 144), 96, 30);
        var btnSalir = CrearBoton("Salir", 574, Color.FromArgb(220, 53, 69), 88, 30);
        btnProcesar.Top = btnSalir.Top = 8;
        panelTop.Controls.Add(btnProcesar);
        panelTop.Controls.Add(btnSalir);

        panelTop.Controls.Add(new Label { Text = "Buscar", Left = 12, Top = 50, Width = 50, Font = new Font("Segoe UI", 9, FontStyle.Bold) });
        _cboCriterio = new ComboBox
        {
            Left = 66,
            Top = 46,
            Width = 140,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cboCriterio.Items.AddRange(["Pedido", "Teléfono", "Cliente", "Empacador", "Motorizado"]);
        _cboCriterio.SelectedIndex = 1;
        panelTop.Controls.Add(_cboCriterio);

        _txtFiltro = new TextBox { Left = 210, Top = 46, Width = 246 };
        panelTop.Controls.Add(_txtFiltro);

        _lblRegistro = new Label
        {
            Left = 470,
            Top = 50,
            Width = 280,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Color.FromArgb(35, 35, 35),
            Text = "Registro 0 de 0"
        };
        panelTop.Controls.Add(_lblRegistro);

        _gridCabecera = new DataGridView
        {
            Dock = DockStyle.Top,
            Height = 300,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            RowHeadersVisible = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = false,
            BackgroundColor = Color.White
        };
        _gridCabecera.Columns.AddRange(
            new DataGridViewTextBoxColumn { HeaderText = "Pedido", DataPropertyName = nameof(PedidoDeliverySeguimiento.CodigoPedido), Width = 90 },
            new DataGridViewTextBoxColumn { HeaderText = "Fecha", DataPropertyName = nameof(PedidoDeliverySeguimiento.FechaRegistro), Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM HH:mm" } },
            new DataGridViewTextBoxColumn { HeaderText = "Usuario", DataPropertyName = nameof(PedidoDeliverySeguimiento.Usuario), Width = 90 },
            new DataGridViewTextBoxColumn { HeaderText = "Caja", DataPropertyName = nameof(PedidoDeliverySeguimiento.Caja), Width = 70 },
            new DataGridViewTextBoxColumn { HeaderText = "Teléfono", DataPropertyName = nameof(PedidoDeliverySeguimiento.Telefono), Width = 115 },
            new DataGridViewTextBoxColumn { HeaderText = "Cliente", DataPropertyName = nameof(PedidoDeliverySeguimiento.Cliente), Width = 180 },
            new DataGridViewTextBoxColumn { HeaderText = "Empacador", DataPropertyName = nameof(PedidoDeliverySeguimiento.Empacador), Width = 115 },
            new DataGridViewTextBoxColumn { HeaderText = "Motorizado", DataPropertyName = nameof(PedidoDeliverySeguimiento.Motorizado), Width = 115 },
            new DataGridViewTextBoxColumn { HeaderText = "H.Asign", DataPropertyName = nameof(PedidoDeliverySeguimiento.FechaAsignacion), Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "HH:mm" } },
            new DataGridViewTextBoxColumn { HeaderText = "H.Salid", DataPropertyName = nameof(PedidoDeliverySeguimiento.FechaSalida), Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "HH:mm" } },
            new DataGridViewTextBoxColumn { HeaderText = "H.Llega", DataPropertyName = nameof(PedidoDeliverySeguimiento.FechaLlegada), Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "HH:mm" } }
        );

        var panelInfo = new Panel
        {
            Dock = DockStyle.Top,
            Height = 66,
            Padding = new Padding(12, 8, 12, 4),
            BackColor = Color.White
        };
        _lblReferencia = new Label { Left = 12, Top = 8, Width = 1120, Font = new Font("Segoe UI", 9), Text = "Referencia: -" };
        _lblDireccion = new Label { Left = 12, Top = 28, Width = 1120, Font = new Font("Segoe UI", 9), Text = "Dirección: -" };
        _lblZona = new Label { Left = 12, Top = 48, Width = 1120, Font = new Font("Segoe UI", 9), Text = "Zona: -" };
        panelInfo.Controls.AddRange([_lblReferencia, _lblDireccion, _lblZona]);

        _gridDetalle = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            RowHeadersVisible = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = false,
            BackgroundColor = Color.White
        };
        _gridDetalle.Columns.AddRange(
            new DataGridViewTextBoxColumn { HeaderText = "Ítem", DataPropertyName = "Item", Width = 70 },
            new DataGridViewTextBoxColumn { HeaderText = "Producto", DataPropertyName = "Producto", Width = 310 },
            new DataGridViewTextBoxColumn { HeaderText = "P.Venta", DataPropertyName = "PrecioUnitario", Width = 95, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { HeaderText = "Cantidad", DataPropertyName = "Cantidad", Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { HeaderText = "SubTotal", DataPropertyName = "SubTotal", Width = 95, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } }
        );

        var panelNavegacion = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 46,
            Padding = new Padding(12, 8, 12, 8),
            BackColor = Color.White
        };
        var btnPrimero = new Button { Text = "|<", Left = 12, Top = 8, Width = 36, Height = 28 };
        var btnPrevio = new Button { Text = "<", Left = 52, Top = 8, Width = 36, Height = 28 };
        var btnSiguiente = new Button { Text = ">", Left = 92, Top = 8, Width = 36, Height = 28 };
        var btnUltimo = new Button { Text = ">|", Left = 132, Top = 8, Width = 36, Height = 28 };
        panelNavegacion.Controls.AddRange([btnPrimero, btnPrevio, btnSiguiente, btnUltimo]);

        Controls.Add(_gridDetalle);
        Controls.Add(panelInfo);
        Controls.Add(_gridCabecera);
        Controls.Add(panelTop);
        Controls.Add(panelNavegacion);

        Load += async (_, _) => await RecargarAsync();
        btnProcesar.Click += async (_, _) => await RecargarAsync();
        btnSalir.Click += (_, _) => Close();
        _txtFiltro.TextChanged += (_, _) => AplicarFiltro();
        _gridCabecera.SelectionChanged += async (_, _) => await MostrarDetalleActualAsync();
        btnPrimero.Click += (_, _) => SeleccionarIndice(0);
        btnPrevio.Click += (_, _) => SeleccionarIndice(Math.Max(0, (_gridCabecera.CurrentRow?.Index ?? 0) - 1));
        btnSiguiente.Click += (_, _) => SeleccionarIndice(Math.Min(_gridCabecera.Rows.Count - 1, (_gridCabecera.CurrentRow?.Index ?? 0) + 1));
        btnUltimo.Click += (_, _) => SeleccionarIndice(_gridCabecera.Rows.Count - 1);
    }

    private async Task RecargarAsync()
    {
        Cursor = Cursors.WaitCursor;
        try
        {
            var result = await _seguimientoHandler.HandleAsync(
                new ObtenerPedidosSeguimientoDeliveryEntregadosQuery(_dtpInicio.Value, _dtpFin.Value));
            if (!result.EsExitoso || result.Valor is null)
            {
                MessageBox.Show(result.MensajeError ?? "No se pudo cargar la lista de pedidos entregados.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _todos = result.Valor;
            _filtrados = _todos.ToList();
            _binding.DataSource = _filtrados;
            _gridCabecera.DataSource = _binding;
            SeleccionarIndice(0);
            ActualizarRegistro();
            await MostrarDetalleActualAsync();
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private void AplicarFiltro()
    {
        var texto = _txtFiltro.Text.Trim();
        if (string.IsNullOrEmpty(texto))
        {
            _filtrados = _todos.ToList();
        }
        else
        {
            _filtrados = _todos.Where(x => ObtenerCampoFiltro(x).Contains(texto, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        _binding.DataSource = _filtrados;
        _gridCabecera.DataSource = _binding;
        SeleccionarIndice(0);
        ActualizarRegistro();
    }

    private string ObtenerCampoFiltro(PedidoDeliverySeguimiento item) => _cboCriterio.Text switch
    {
        "Pedido" => item.CodigoPedido,
        "Cliente" => item.Cliente ?? string.Empty,
        "Empacador" => item.Empacador ?? string.Empty,
        "Motorizado" => item.Motorizado ?? string.Empty,
        _ => item.Telefono ?? string.Empty
    };

    private async Task MostrarDetalleActualAsync()
    {
        if (ObtenerSeleccionado() is not { } pedido)
        {
            _gridDetalle.DataSource = null;
            _lblReferencia.Text = "Referencia: -";
            _lblDireccion.Text = "Dirección: -";
            _lblZona.Text = "Zona: -";
            return;
        }

        _lblReferencia.Text = $"Referencia: {pedido.Referencia ?? string.Empty}";
        _lblDireccion.Text = $"Dirección: {pedido.Direccion ?? string.Empty}";
        _lblZona.Text = $"Zona: {pedido.Zona ?? string.Empty}";

        var detalle = await _detalleHandler.HandleAsync(new ObtenerDetallePedidoExtendidoQuery(pedido.CodigoPedido));
        _gridDetalle.DataSource = detalle.EsExitoso && detalle.Valor is not null
            ? detalle.Valor.Items.ToList()
            : null;
    }

    private PedidoDeliverySeguimiento? ObtenerSeleccionado()
        => _gridCabecera.CurrentRow?.DataBoundItem as PedidoDeliverySeguimiento;

    private void SeleccionarIndice(int index)
    {
        if (_gridCabecera.Rows.Count == 0 || index < 0 || index >= _gridCabecera.Rows.Count)
            return;

        _gridCabecera.ClearSelection();
        _gridCabecera.Rows[index].Selected = true;
        _gridCabecera.CurrentCell = _gridCabecera.Rows[index].Cells[0];
        ActualizarRegistro();
    }

    private void ActualizarRegistro()
    {
        var actual = _gridCabecera.Rows.Count == 0 ? 0 : (_gridCabecera.CurrentRow?.Index ?? 0) + 1;
        _lblRegistro.Text = $"Registro {actual} de {_gridCabecera.Rows.Count}";
    }

    private static Button CrearBoton(string texto, int left, Color color, int width, int height)
    {
        var button = new Button
        {
            Text = texto,
            Left = left,
            Width = width,
            Height = height,
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        button.FlatAppearance.BorderSize = 0;
        return button;
    }
}
