using System.Net;
using Inforest.Application.Delivery;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario de mantenimiento/listado principal de clientes delivery frecuentes.
/// Legacy: <c>frmClienteDelivery.frm</c>.
/// Reglas: BR-DEL-037-001, BR-DEL-037-002, BR-DEL-037-003, BR-DEL-037-004.
/// </summary>
public sealed class FrmClienteDelivery : Form
{
    private readonly ObtenerClientesDeliveryListadoHandler _listarHandler;
    private readonly ObtenerSiguienteCodigoClienteDeliveryHandler _codigoClienteHandler;
    private readonly ObtenerClienteDeliveryPorCodigoHandler _clientePorCodigoHandler;
    private readonly CrearClienteDeliveryHandler _crearHandler;
    private readonly ActualizarClienteDeliveryHandler _actualizarHandler;
    private readonly ActualizarFotoClienteDeliveryHandler _actualizarFotoHandler;

    private readonly BindingSource _bindingSource = new();
    private IReadOnlyList<ClienteDeliveryListadoItem> _todos = [];
    private List<ClienteDeliveryListadoItem> _filtrados = [];
    private string _columnaOrden = nameof(ClienteDeliveryListadoItem.Cliente);
    private bool _ordenAscendente = true;

    private DataGridView _grid = null!;
    private ComboBox _cmbFiltro = null!;
    private TextBox _txtFiltro = null!;
    private Label _lblRegistro = null!;
    private Button _btnPrimero = null!;
    private Button _btnAnterior = null!;
    private Button _btnSiguiente = null!;
    private Button _btnUltimo = null!;

    public FrmClienteDelivery(
        ObtenerClientesDeliveryListadoHandler listarHandler,
        ObtenerSiguienteCodigoClienteDeliveryHandler codigoClienteHandler,
        ObtenerClienteDeliveryPorCodigoHandler clientePorCodigoHandler,
        CrearClienteDeliveryHandler crearHandler,
        ActualizarClienteDeliveryHandler actualizarHandler,
        ActualizarFotoClienteDeliveryHandler actualizarFotoHandler)
    {
        _listarHandler = listarHandler;
        _codigoClienteHandler = codigoClienteHandler;
        _clientePorCodigoHandler = clientePorCodigoHandler;
        _crearHandler = crearHandler;
        _actualizarHandler = actualizarHandler;
        _actualizarFotoHandler = actualizarFotoHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Mantenimiento de Clientes Frecuentes";
        WindowState = FormWindowState.Maximized;
        MinimumSize = new Size(1100, 640);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(245, 246, 247);

        var filtros = new Panel
        {
            Dock = DockStyle.Top,
            Height = 52,
            Padding = new Padding(16, 12, 16, 8),
            BackColor = Color.White
        };

        var lblFiltro = new Label { Text = "Filtrar por:", Left = 12, Top = 16, Width = 72 };
        _cmbFiltro = new ComboBox
        {
            Left = 90,
            Top = 12,
            Width = 160,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbFiltro.Items.AddRange(["Código", "Tipo", "Cliente", "Teléfono", "Dirección", "Estado", "Activo"]);
        _cmbFiltro.SelectedIndex = 2;

        _txtFiltro = new TextBox { Left = 260, Top = 12, Width = 320 };
        var btnLimpiar = new Button
        {
            Text = "Limpiar",
            Left = 590,
            Top = 11,
            Width = 90,
            Height = 28
        };
        btnLimpiar.Click += (_, _) => _txtFiltro.Clear();
        filtros.Controls.AddRange([lblFiltro, _cmbFiltro, _txtFiltro, btnLimpiar]);

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            RowHeadersVisible = false,
            AutoGenerateColumns = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        };
        _grid.Columns.AddRange(
            new DataGridViewTextBoxColumn { Name = "colCodigo", HeaderText = "Código", DataPropertyName = nameof(ClienteDeliveryListadoItem.Codigo), Width = 90 },
            new DataGridViewTextBoxColumn { Name = "colTipo", HeaderText = "Tipo", DataPropertyName = nameof(ClienteDeliveryListadoItem.TipoCliente), Width = 120 },
            new DataGridViewTextBoxColumn { Name = "colCliente", HeaderText = "Cliente", DataPropertyName = nameof(ClienteDeliveryListadoItem.Cliente), Width = 250 },
            new DataGridViewTextBoxColumn { Name = "colTelefono", HeaderText = "Teléfono", DataPropertyName = nameof(ClienteDeliveryListadoItem.Telefono), Width = 120 },
            new DataGridViewTextBoxColumn { Name = "colDireccion", HeaderText = "Dirección", DataPropertyName = nameof(ClienteDeliveryListadoItem.Direccion), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill },
            new DataGridViewTextBoxColumn { Name = "colLinea", HeaderText = "Línea", DataPropertyName = nameof(ClienteDeliveryListadoItem.Linea), Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { Name = "colDescuento", HeaderText = "Descuento", DataPropertyName = nameof(ClienteDeliveryListadoItem.Descuento), Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { Name = "colEstado", HeaderText = "Estado", DataPropertyName = nameof(ClienteDeliveryListadoItem.EstadoFrecuente), Width = 130 },
            new DataGridViewCheckBoxColumn { Name = "colActivo", HeaderText = "Activo", DataPropertyName = nameof(ClienteDeliveryListadoItem.Activo), Width = 70 }
        );

        var pie = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 58,
            Padding = new Padding(12, 8, 12, 8),
            BackColor = Color.White
        };

        _btnPrimero = CrearBotonAccion("|<", 12);
        _btnAnterior = CrearBotonAccion("<", 56);
        _btnSiguiente = CrearBotonAccion(">", 100);
        _btnUltimo = CrearBotonAccion(">|", 144);
        _lblRegistro = new Label
        {
            Left = 196,
            Top = 18,
            Width = 180,
            Text = "Registro 0 de 0",
            Font = new Font("Segoe UI", 9F, FontStyle.Bold)
        };

        var btnNuevo = CrearBotonPrincipal("Nuevo", 430, Color.FromArgb(37, 99, 235));
        var btnModificar = CrearBotonPrincipal("Modificar", 538, Color.FromArgb(14, 116, 144));
        var btnVistaPrevia = CrearBotonPrincipal("Vista previa", 646, Color.FromArgb(22, 163, 74));
        var btnExportar = CrearBotonPrincipal("Exportar HTML", 786, Color.FromArgb(107, 114, 128));
        var btnSalir = CrearBotonPrincipal("Salir", 926, Color.FromArgb(220, 53, 69));

        pie.Controls.AddRange([
            _btnPrimero, _btnAnterior, _btnSiguiente, _btnUltimo, _lblRegistro,
            btnNuevo, btnModificar, btnVistaPrevia, btnExportar, btnSalir
        ]);

        Controls.Add(_grid);
        Controls.Add(pie);
        Controls.Add(filtros);

        _txtFiltro.TextChanged += (_, _) => AplicarFiltro();
        _cmbFiltro.SelectedIndexChanged += (_, _) => AplicarFiltro();
        _grid.CellDoubleClick += async (_, _) => await AbrirDetalleAsync(false);
        _grid.SelectionChanged += (_, _) => ActualizarRegistro();
        _grid.ColumnHeaderMouseClick += (_, e) => OrdenarPorColumna(e.ColumnIndex);
        _btnPrimero.Click += (_, _) => MoverPrimero();
        _btnAnterior.Click += (_, _) => MoverAnterior();
        _btnSiguiente.Click += (_, _) => MoverSiguiente();
        _btnUltimo.Click += (_, _) => MoverUltimo();
        btnNuevo.Click += async (_, _) => await AbrirDetalleAsync(true);
        btnModificar.Click += async (_, _) => await AbrirDetalleAsync(false);
        btnVistaPrevia.Click += (_, _) => MostrarVistaPrevia();
        btnExportar.Click += async (_, _) => await ExportarHtmlAsync();
        btnSalir.Click += (_, _) => Close();
        Load += async (_, _) => await CargarAsync();
    }

    private Button CrearBotonAccion(string texto, int left)
        => new()
        {
            Text = texto,
            Left = left,
            Top = 12,
            Width = 36,
            Height = 32
        };

    private static Button CrearBotonPrincipal(string texto, int left, Color color)
    {
        var boton = new Button
        {
            Text = texto,
            Left = left,
            Top = 10,
            Width = texto == "Exportar HTML" ? 132 : 100,
            Height = 34,
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        boton.FlatAppearance.BorderSize = 0;
        return boton;
    }

    private async Task CargarAsync()
    {
        var result = await _listarHandler.HandleAsync(new ObtenerClientesDeliveryListadoQuery());
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError ?? "No se pudo cargar la lista de clientes delivery.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _todos = result.Valor ?? [];
        _filtrados = _todos.ToList();
        AplicarOrden();
    }

    private void AplicarFiltro()
    {
        var texto = _txtFiltro.Text.Trim();
        if (string.IsNullOrWhiteSpace(texto))
        {
            _filtrados = _todos.ToList();
            AplicarOrden();
            return;
        }

        _filtrados = _todos.Where(item => ObtenerValorFiltro(item).Contains(texto, StringComparison.OrdinalIgnoreCase)).ToList();
        AplicarOrden();
    }

    private string ObtenerValorFiltro(ClienteDeliveryListadoItem item)
        => _cmbFiltro.SelectedIndex switch
        {
            0 => item.Codigo,
            1 => item.TipoCliente ?? string.Empty,
            2 => item.Cliente ?? string.Empty,
            3 => item.Telefono ?? string.Empty,
            4 => item.Direccion ?? string.Empty,
            5 => item.EstadoFrecuente ?? string.Empty,
            6 => item.Activo ? "SI" : "NO",
            _ => item.Cliente ?? string.Empty
        };

    private void OrdenarPorColumna(int columnIndex)
    {
        var property = columnIndex switch
        {
            0 => nameof(ClienteDeliveryListadoItem.Codigo),
            1 => nameof(ClienteDeliveryListadoItem.TipoCliente),
            2 => nameof(ClienteDeliveryListadoItem.Cliente),
            3 => nameof(ClienteDeliveryListadoItem.Telefono),
            4 => nameof(ClienteDeliveryListadoItem.Direccion),
            5 => nameof(ClienteDeliveryListadoItem.Linea),
            6 => nameof(ClienteDeliveryListadoItem.Descuento),
            7 => nameof(ClienteDeliveryListadoItem.EstadoFrecuente),
            8 => nameof(ClienteDeliveryListadoItem.Activo),
            _ => nameof(ClienteDeliveryListadoItem.Cliente)
        };

        _ordenAscendente = _columnaOrden == property ? !_ordenAscendente : true;
        _columnaOrden = property;
        AplicarOrden();
    }

    private void AplicarOrden()
    {
        IEnumerable<ClienteDeliveryListadoItem> query = _columnaOrden switch
        {
            nameof(ClienteDeliveryListadoItem.Codigo) => _ordenAscendente ? _filtrados.OrderBy(x => x.Codigo) : _filtrados.OrderByDescending(x => x.Codigo),
            nameof(ClienteDeliveryListadoItem.TipoCliente) => _ordenAscendente ? _filtrados.OrderBy(x => x.TipoCliente) : _filtrados.OrderByDescending(x => x.TipoCliente),
            nameof(ClienteDeliveryListadoItem.Telefono) => _ordenAscendente ? _filtrados.OrderBy(x => x.Telefono) : _filtrados.OrderByDescending(x => x.Telefono),
            nameof(ClienteDeliveryListadoItem.Direccion) => _ordenAscendente ? _filtrados.OrderBy(x => x.Direccion) : _filtrados.OrderByDescending(x => x.Direccion),
            nameof(ClienteDeliveryListadoItem.Linea) => _ordenAscendente ? _filtrados.OrderBy(x => x.Linea) : _filtrados.OrderByDescending(x => x.Linea),
            nameof(ClienteDeliveryListadoItem.Descuento) => _ordenAscendente ? _filtrados.OrderBy(x => x.Descuento) : _filtrados.OrderByDescending(x => x.Descuento),
            nameof(ClienteDeliveryListadoItem.EstadoFrecuente) => _ordenAscendente ? _filtrados.OrderBy(x => x.EstadoFrecuente) : _filtrados.OrderByDescending(x => x.EstadoFrecuente),
            nameof(ClienteDeliveryListadoItem.Activo) => _ordenAscendente ? _filtrados.OrderBy(x => x.Activo) : _filtrados.OrderByDescending(x => x.Activo),
            _ => _ordenAscendente ? _filtrados.OrderBy(x => x.Cliente) : _filtrados.OrderByDescending(x => x.Cliente)
        };

        _bindingSource.DataSource = query.ToList();
        _grid.DataSource = _bindingSource;
        if (_grid.Rows.Count > 0)
            _grid.Rows[0].Selected = true;
        ActualizarRegistro();
    }

    private async Task AbrirDetalleAsync(bool esNuevo)
    {
        var codigoEdicion = esNuevo ? null : ObtenerCodigoSeleccionado();
        if (!esNuevo && string.IsNullOrWhiteSpace(codigoEdicion))
        {
            MessageBox.Show("No existe datos ingresados.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dlg = new NuevoDeliveryForm(
            _codigoClienteHandler,
            _clientePorCodigoHandler,
            _crearHandler,
            _actualizarHandler,
            _actualizarFotoHandler,
            codigoEdicion);
        dlg.ShowDialog(this);
        await CargarAsync();
        SeleccionarFila(codigoEdicion ?? dlg.ClienteResultado?.CodigoDelivery);
    }

    private void MostrarVistaPrevia()
    {
        if (_grid.Rows.Count == 0)
            return;

        using var preview = new Form
        {
            Text = "Vista previa — Clientes delivery",
            Size = new Size(980, 560),
            StartPosition = FormStartPosition.CenterParent
        };

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoGenerateColumns = false,
            DataSource = _bindingSource.DataSource
        };

        foreach (DataGridViewColumn column in _grid.Columns)
        {
            var clone = (DataGridViewColumn)column.Clone();
            clone.DataPropertyName = column.DataPropertyName;
            clone.HeaderText = column.HeaderText;
            grid.Columns.Add(clone);
        }

        preview.Controls.Add(grid);
        preview.ShowDialog(this);
    }

    private async Task ExportarHtmlAsync()
    {
        if (_filtrados.Count == 0)
            return;

        using var dialog = new SaveFileDialog
        {
            Filter = "Hyper Text Markup Language|*.html",
            FileName = "clientes-delivery.html"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        var html = """
            <html>
            <head>
                <meta charset="utf-8" />
                <title>Clientes Delivery</title>
                <style>
                    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; }
                    table { border-collapse: collapse; width: 100%; }
                    th, td { border: 1px solid #d0d7de; padding: 8px; font-size: 12px; }
                    th { background: #f3f4f6; text-align: left; }
                    td.num { text-align: right; }
                </style>
            </head>
            <body>
                <h2>Clientes Delivery</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Código</th><th>Tipo</th><th>Cliente</th><th>Teléfono</th><th>Dirección</th>
                            <th>Línea</th><th>Descuento</th><th>Estado</th><th>Activo</th>
                        </tr>
                    </thead>
                    <tbody>
            """;

        var rows = string.Join(Environment.NewLine, _filtrados.Select(item =>
            $"<tr><td>{WebUtility.HtmlEncode(item.Codigo)}</td><td>{WebUtility.HtmlEncode(item.TipoCliente ?? string.Empty)}</td><td>{WebUtility.HtmlEncode(item.Cliente ?? string.Empty)}</td><td>{WebUtility.HtmlEncode(item.Telefono ?? string.Empty)}</td><td>{WebUtility.HtmlEncode(item.Direccion ?? string.Empty)}</td><td class=\"num\">{item.Linea:N2}</td><td class=\"num\">{item.Descuento:N2}</td><td>{WebUtility.HtmlEncode(item.EstadoFrecuente ?? string.Empty)}</td><td>{(item.Activo ? "Sí" : "No")}</td></tr>"));

        var closing = """
                    </tbody>
                </table>
            </body>
            </html>
            """;

        await File.WriteAllTextAsync(dialog.FileName, html + rows + closing);
        MessageBox.Show("Archivo HTML exportado correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void MoverPrimero()
    {
        if (_grid.Rows.Count == 0) return;
        _grid.ClearSelection();
        _grid.Rows[0].Selected = true;
        _grid.CurrentCell = _grid.Rows[0].Cells[0];
        ActualizarRegistro();
    }

    private void MoverAnterior()
    {
        if (_grid.CurrentRow is null || _grid.CurrentRow.Index <= 0) return;
        var index = _grid.CurrentRow.Index - 1;
        SeleccionarIndice(index);
    }

    private void MoverSiguiente()
    {
        if (_grid.CurrentRow is null || _grid.CurrentRow.Index >= _grid.Rows.Count - 1) return;
        var index = _grid.CurrentRow.Index + 1;
        SeleccionarIndice(index);
    }

    private void MoverUltimo()
    {
        if (_grid.Rows.Count == 0) return;
        SeleccionarIndice(_grid.Rows.Count - 1);
    }

    private void SeleccionarIndice(int index)
    {
        _grid.ClearSelection();
        _grid.Rows[index].Selected = true;
        _grid.CurrentCell = _grid.Rows[index].Cells[0];
        ActualizarRegistro();
    }

    private void SeleccionarFila(string? codigo)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            return;

        foreach (DataGridViewRow row in _grid.Rows)
        {
            if (string.Equals(row.Cells["colCodigo"].Value?.ToString(), codigo, StringComparison.OrdinalIgnoreCase))
            {
                _grid.ClearSelection();
                row.Selected = true;
                _grid.CurrentCell = row.Cells[0];
                break;
            }
        }

        ActualizarRegistro();
    }

    private string? ObtenerCodigoSeleccionado()
        => _grid.CurrentRow?.Cells["colCodigo"].Value?.ToString();

    private void ActualizarRegistro()
    {
        var total = _grid.Rows.Count;
        var actual = total == 0 || _grid.CurrentRow is null ? 0 : _grid.CurrentRow.Index + 1;
        _lblRegistro.Text = $"Registro {actual} de {total}";
    }
}
