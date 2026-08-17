using System.Drawing.Printing;
using Inforest.Application.Delivery;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Desktop.Clientes;

/// <summary>
/// Mantenimiento WinForms de tarjetas RFID/proximidad.
/// Legacy: <c>FrmTarjetaAproximidad.frm</c> + <c>FrmTarjetaAproximidadDetalle.frm</c>.
/// Reglas: BR-RFID-001, BR-RFID-002, BR-RFID-003, BR-RFID-004.
/// </summary>
public sealed class FrmTarjetaProximidad : Form
{
    private static readonly Color SurfaceColor = Color.FromArgb(245, 246, 247);
    private static readonly Color HeaderColor = Color.FromArgb(235, 236, 240);
    private static readonly Color AccentColor = Color.FromArgb(0, 122, 204);

    private readonly ObtenerTarjetasProximidadHandler _obtenerTarjetasHandler;
    private readonly CrearTarjetaProximidadHandler _crearHandler;
    private readonly ActualizarTarjetaProximidadHandler _actualizarHandler;
    private readonly ObtenerMovimientosTarjetaProximidadHandler _obtenerMovimientosHandler;
    private readonly IClienteDeliveryRepository _clienteDeliveryRepository;

    private readonly DataGridView _dgvTarjetas = CreateGrid();
    private readonly DataGridView _dgvMovimientos = CreateGrid();
    private readonly TextBox _txtCodigo = new() { Width = 180 };
    private readonly TextBox _txtDescripcion = new() { Width = 260 };
    private readonly TextBox _txtCodigoCliente = new() { Width = 120, ReadOnly = true };
    private readonly TextBox _txtCliente = new() { Width = 260, ReadOnly = true };
    private readonly ComboBox _cbEstado = new() { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly NumericUpDown _nudMonto = new()
    {
        Width = 140,
        DecimalPlaces = 2,
        Maximum = 999999999,
        ThousandsSeparator = true
    };
    private readonly Label _lblEstado = new()
    {
        AutoSize = true,
        ForeColor = Color.FromArgb(35, 35, 35)
    };

    private List<TarjetaProximidad> _tarjetas = [];
    private List<MovimientoTarjetaProximidad> _movimientos = [];
    private TarjetaProximidad? _tarjetaSeleccionada;
    private bool _modoNuevo;
    private int _lineaImpresion;

    public FrmTarjetaProximidad(
        ObtenerTarjetasProximidadHandler obtenerTarjetasHandler,
        CrearTarjetaProximidadHandler crearHandler,
        ActualizarTarjetaProximidadHandler actualizarHandler,
        ObtenerMovimientosTarjetaProximidadHandler obtenerMovimientosHandler,
        IClienteDeliveryRepository clienteDeliveryRepository)
    {
        _obtenerTarjetasHandler = obtenerTarjetasHandler;
        _crearHandler = crearHandler;
        _actualizarHandler = actualizarHandler;
        _obtenerMovimientosHandler = obtenerMovimientosHandler;
        _clienteDeliveryRepository = clienteDeliveryRepository;

        Text = "Tarjetas de proximidad";
        Width = 1360;
        Height = 760;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = SurfaceColor;

        _cbEstado.Items.AddRange(["Free", "Asignada", "Bloqueado"]);
        _cbEstado.SelectedIndex = 0;

        Controls.Add(BuildLayout());

        Load += async (_, _) => await CargarAsync();
        _dgvTarjetas.SelectionChanged += async (_, _) => await CargarSeleccionAsync();
    }

    private Control BuildLayout()
    {
        var header = new Panel
        {
            Dock = DockStyle.Top,
            Height = 64,
            Padding = new Padding(18, 14, 18, 14),
            BackColor = HeaderColor
        };
        header.Controls.Add(new Label
        {
            Dock = DockStyle.Left,
            AutoSize = false,
            Width = 320,
            Text = "Mantenimiento de tarjetas RFID",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.FromArgb(35, 35, 35),
            TextAlign = ContentAlignment.MiddleLeft
        });
        _lblEstado.Dock = DockStyle.Right;
        _lblEstado.Padding = new Padding(0, 8, 0, 0);
        header.Controls.Add(_lblEstado);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 54,
            Padding = new Padding(18, 10, 18, 10),
            BackColor = SurfaceColor
        };

        var btnNuevo = CreateButton("Nuevo", AccentColor, (_, _) => PrepararNuevo());
        var btnGuardar = CreateButton("Guardar", AccentColor, async (_, _) => await GuardarAsync());
        var btnBuscarCliente = CreateButton("Cliente…", Color.FromArgb(108, 117, 125), async (_, _) => await BuscarClienteAsync());
        var btnExportar = CreateButton("Exportar CSV", Color.FromArgb(32, 156, 98), (_, _) => ExportarCsv());
        var btnImprimir = CreateButton("Imprimir movimientos", Color.FromArgb(111, 66, 193), (_, _) => ImprimirMovimientos());
        var btnSalir = CreateButton("Cerrar", Color.FromArgb(220, 53, 69), (_, _) => Close());

        actions.Controls.AddRange([btnNuevo, btnGuardar, btnBuscarCliente, btnExportar, btnImprimir, btnSalir]);

        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            SplitterDistance = 690,
            BackColor = SurfaceColor,
            Panel1MinSize = 480,
            Panel2MinSize = 420
        };

        split.Panel1.Padding = new Padding(18, 0, 9, 18);
        split.Panel2.Padding = new Padding(9, 0, 18, 18);
        split.Panel1.Controls.Add(CreateCardPanel("Tarjetas registradas", _dgvTarjetas));
        split.Panel2.Controls.Add(BuildRightPanel());

        var root = new Panel { Dock = DockStyle.Fill };
        root.Controls.Add(split);
        root.Controls.Add(actions);
        root.Controls.Add(header);
        return root;
    }

    private Control BuildRightPanel()
    {
        var right = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        right.RowStyles.Add(new RowStyle(SizeType.Absolute, 290));
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        right.Controls.Add(BuildEditorCard(), 0, 0);
        right.Controls.Add(CreateCardPanel("Últimos 10 movimientos", _dgvMovimientos), 0, 1);
        return right;
    }

    private Control BuildEditorCard()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(18),
            BorderStyle = BorderStyle.FixedSingle
        };

        panel.Controls.Add(new Label
        {
            Dock = DockStyle.Top,
            Height = 28,
            Text = "Datos de la tarjeta",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(35, 35, 35)
        });

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            Padding = new Padding(0, 12, 0, 0)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        AddField(layout, 0, "Código RFID", _txtCodigo);
        AddField(layout, 1, "Descripción", _txtDescripcion);
        AddField(layout, 2, "Estado", _cbEstado);
        AddField(layout, 3, "Saldo", _nudMonto);

        var clientPanel = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            Margin = new Padding(0)
        };
        clientPanel.Controls.AddRange([_txtCodigoCliente, _txtCliente]);
        AddField(layout, 4, "Cliente delivery", clientPanel);

        panel.Controls.Add(layout);
        return panel;
    }

    private async Task CargarAsync()
    {
        var result = await _obtenerTarjetasHandler.HandleAsync(new ObtenerTarjetasProximidadQuery());
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _tarjetas = result.Valor!.ToList();
        _dgvTarjetas.DataSource = null;
        _dgvTarjetas.DataSource = _tarjetas
            .Select(t => new
            {
                t.CodigoRfid,
                t.Descripcion,
                t.MontoDisponible,
                t.CodigoCliente,
                t.NombreCliente,
                t.Estado
            })
            .ToList();

        ConfigurarColumnasTarjetas();
        _lblEstado.Text = $"Registros: {_tarjetas.Count}";

        if (_tarjetas.Count == 0)
        {
            PrepararNuevo();
            _dgvMovimientos.DataSource = null;
            return;
        }

        _dgvTarjetas.Rows[0].Selected = true;
        await CargarSeleccionAsync();
    }

    private async Task CargarSeleccionAsync()
    {
        if (_modoNuevo)
            return;

        var index = _dgvTarjetas.CurrentCell?.RowIndex ?? -1;
        if (index < 0 || index >= _tarjetas.Count)
            return;

        _tarjetaSeleccionada = _tarjetas[index];
        _txtCodigo.Enabled = false;
        _txtCodigo.Text = _tarjetaSeleccionada.CodigoRfid;
        _txtDescripcion.Text = _tarjetaSeleccionada.Descripcion;
        _txtCodigoCliente.Text = _tarjetaSeleccionada.CodigoCliente;
        _txtCliente.Text = _tarjetaSeleccionada.NombreCliente;
        _cbEstado.SelectedItem = _tarjetaSeleccionada.Estado;
        _nudMonto.Value = Math.Min(_nudMonto.Maximum, _tarjetaSeleccionada.MontoDisponible);

        await CargarMovimientosAsync(_tarjetaSeleccionada.CodigoRfid);
    }

    private async Task CargarMovimientosAsync(string codigoRfid)
    {
        var result = await _obtenerMovimientosHandler.HandleAsync(new ObtenerMovimientosTarjetaProximidadQuery(codigoRfid));
        _movimientos = result.EsExitoso ? result.Valor!.ToList() : [];
        _dgvMovimientos.DataSource = null;
        _dgvMovimientos.DataSource = _movimientos
            .Select(m => new
            {
                Fecha = m.FechaRegistro.ToString("dd/MM/yyyy"),
                m.MontoIngreso,
                m.MontoSalida,
                m.MontoAnterior,
                m.MontoFinal
            })
            .ToList();
        ConfigurarColumnasMovimientos();
    }

    private void PrepararNuevo()
    {
        _modoNuevo = true;
        _tarjetaSeleccionada = null;
        _txtCodigo.Enabled = true;
        _txtCodigo.Clear();
        _txtDescripcion.Clear();
        _txtCodigoCliente.Clear();
        _txtCliente.Clear();
        _cbEstado.SelectedItem = "Free";
        _nudMonto.Value = 0;
        _dgvMovimientos.DataSource = null;
        _movimientos = [];
        _txtCodigo.Focus();
    }

    private async Task GuardarAsync()
    {
        if (string.IsNullOrWhiteSpace(_txtCodigoCliente.Text))
        {
            MessageBox.Show("Debe seleccionar un cliente delivery.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var usuario = Environment.UserName;
        Result result;

        if (_modoNuevo || _tarjetaSeleccionada is null || _txtCodigo.Enabled)
        {
            result = await _crearHandler.HandleAsync(new CrearTarjetaProximidadCommand(
                _txtCodigo.Text,
                _txtDescripcion.Text,
                _nudMonto.Value,
                _txtCodigoCliente.Text,
                _cbEstado.Text,
                usuario));
        }
        else
        {
            result = await _actualizarHandler.HandleAsync(new ActualizarTarjetaProximidadCommand(
                _txtCodigo.Text,
                _txtDescripcion.Text,
                _nudMonto.Value,
                _txtCodigoCliente.Text,
                _cbEstado.Text,
                usuario));
        }

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _modoNuevo = false;
        await CargarAsync();
        SeleccionarFila(_txtCodigo.Text.Trim());
    }

    private async Task BuscarClienteAsync()
    {
        using var dialog = new SeleccionarClienteDeliveryDialog(_clienteDeliveryRepository);
        if (dialog.ShowDialog(this) != DialogResult.OK || dialog.ClienteSeleccionado is null)
            return;

        _txtCodigoCliente.Text = dialog.ClienteSeleccionado.CodigoDelivery;
        _txtCliente.Text = $"{dialog.ClienteSeleccionado.Apellido} {dialog.ClienteSeleccionado.Nombre}".Trim();
        await Task.CompletedTask;
    }

    private void ExportarCsv()
    {
        if (_tarjetas.Count == 0)
        {
            MessageBox.Show("No hay tarjetas para exportar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = "tarjetas-proximidad.csv"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        using var writer = new StreamWriter(dialog.FileName, false);
        writer.WriteLine("CodigoRfid,Descripcion,MontoDisponible,CodigoCliente,NombreCliente,Estado");
        foreach (var item in _tarjetas)
        {
            writer.WriteLine(
                string.Join(",",
                    Csv(item.CodigoRfid),
                    Csv(item.Descripcion),
                    item.MontoDisponible.ToString("0.00"),
                    Csv(item.CodigoCliente),
                    Csv(item.NombreCliente),
                    Csv(item.Estado)));
        }
    }

    private void ImprimirMovimientos()
    {
        if (_tarjetaSeleccionada is null)
        {
            MessageBox.Show("Seleccione una tarjeta.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (_movimientos.Count == 0)
        {
            MessageBox.Show("La tarjeta no tiene movimientos para imprimir.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var printDocument = new PrintDocument();
        printDocument.DocumentName = $"Movimientos-{_tarjetaSeleccionada.CodigoRfid}";
        printDocument.PrintPage += PrintDocument_PrintPage;
        _lineaImpresion = 0;

        using var preview = new PrintPreviewDialog
        {
            Document = printDocument,
            Width = 960,
            Height = 720
        };
        preview.ShowDialog(this);
    }

    private void PrintDocument_PrintPage(object? sender, PrintPageEventArgs e)
    {
        if (_tarjetaSeleccionada is null || e.Graphics is null)
            return;

        var titleFont = new Font("Segoe UI", 12, FontStyle.Bold);
        var textFont = new Font("Consolas", 9);
        var brush = Brushes.Black;
        float top = e.MarginBounds.Top;

        if (_lineaImpresion == 0)
        {
            e.Graphics.DrawString("Movimientos de tarjeta", titleFont, brush, e.MarginBounds.Left, top);
            top += 28;
            e.Graphics.DrawString($"Cliente: {_tarjetaSeleccionada.NombreCliente}", textFont, brush, e.MarginBounds.Left, top);
            top += 18;
            e.Graphics.DrawString($"Tarjeta: {_tarjetaSeleccionada.CodigoRfid}", textFont, brush, e.MarginBounds.Left, top);
            top += 18;
            e.Graphics.DrawString($"Fecha: {DateTime.Now:dd/MM/yyyy HH:mm}", textFont, brush, e.MarginBounds.Left, top);
            top += 24;
            e.Graphics.DrawString("Fecha        Ingreso    Salida     Anterior   Final", textFont, brush, e.MarginBounds.Left, top);
            top += 18;
        }

        while (_lineaImpresion < _movimientos.Count)
        {
            var mov = _movimientos[_lineaImpresion];
            var line = $"{mov.FechaRegistro:dd/MM/yyyy}  {mov.MontoIngreso,8:0.00}  {mov.MontoSalida,8:0.00}  {mov.MontoAnterior,9:0.00}  {mov.MontoFinal,8:0.00}";
            e.Graphics.DrawString(line, textFont, brush, e.MarginBounds.Left, top);
            top += 18;
            _lineaImpresion++;

            if (top > e.MarginBounds.Bottom - 30)
            {
                e.HasMorePages = true;
                return;
            }
        }

        e.HasMorePages = false;
        _lineaImpresion = 0;
    }

    private void SeleccionarFila(string codigoRfid)
    {
        for (var i = 0; i < _tarjetas.Count; i++)
        {
            if (!string.Equals(_tarjetas[i].CodigoRfid, codigoRfid, StringComparison.OrdinalIgnoreCase))
                continue;

            if (_dgvTarjetas.Rows.Count > i)
            {
                _dgvTarjetas.ClearSelection();
                _dgvTarjetas.Rows[i].Selected = true;
                _dgvTarjetas.CurrentCell = _dgvTarjetas.Rows[i].Cells[0];
            }

            break;
        }
    }

    private void ConfigurarColumnasTarjetas()
    {
        if (_dgvTarjetas.Columns.Count == 0)
            return;

        _dgvTarjetas.Columns[0].HeaderText = "Código";
        _dgvTarjetas.Columns[1].HeaderText = "Descripción";
        _dgvTarjetas.Columns[2].HeaderText = "Saldo";
        _dgvTarjetas.Columns[3].HeaderText = "Cliente";
        _dgvTarjetas.Columns[4].HeaderText = "Nombre cliente";
        _dgvTarjetas.Columns[5].HeaderText = "Estado";
        _dgvTarjetas.Columns[2].DefaultCellStyle.Format = "N2";
        _dgvTarjetas.AutoResizeColumns();
    }

    private void ConfigurarColumnasMovimientos()
    {
        if (_dgvMovimientos.Columns.Count == 0)
            return;

        _dgvMovimientos.Columns[0].HeaderText = "Fecha";
        _dgvMovimientos.Columns[1].HeaderText = "Ingreso";
        _dgvMovimientos.Columns[2].HeaderText = "Salida";
        _dgvMovimientos.Columns[3].HeaderText = "Anterior";
        _dgvMovimientos.Columns[4].HeaderText = "Final";
        for (var i = 1; i < _dgvMovimientos.Columns.Count; i++)
            _dgvMovimientos.Columns[i].DefaultCellStyle.Format = "N2";

        _dgvMovimientos.AutoResizeColumns();
    }

    private static DataGridView CreateGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            MultiSelect = false,
            AutoGenerateColumns = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            RowHeadersVisible = false
        };

        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersDefaultCellStyle.BackColor = HeaderColor;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(234, 245, 255);
        grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(35, 35, 35);
        return grid;
    }

    private static Control CreateCardPanel(string title, Control content)
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };

        panel.Controls.Add(content);
        panel.Controls.Add(new Label
        {
            Dock = DockStyle.Top,
            Height = 34,
            Text = title,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(35, 35, 35),
            Padding = new Padding(12, 8, 0, 0),
            BackColor = Color.White
        });

        content.Dock = DockStyle.Fill;
        return panel;
    }

    private static Button CreateButton(string text, Color color, EventHandler onClick)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            Height = 32,
            Padding = new Padding(12, 0, 12, 0),
            Margin = new Padding(0, 0, 8, 0),
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        button.FlatAppearance.BorderSize = 0;
        button.Click += onClick;
        return button;
    }

    private static void AddField(TableLayoutPanel layout, int row, string label, Control control)
    {
        while (layout.RowCount <= row)
        {
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowCount++;
        }

        layout.Controls.Add(new Label
        {
            Text = label,
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        }, 0, row);

        control.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        layout.Controls.Add(control, 1, row);
    }

    private static string Csv(string? value)
        => $"\"{(value ?? string.Empty).Replace("\"", "\"\"")}\"";

    private sealed class SeleccionarClienteDeliveryDialog : Form
    {
        private readonly IClienteDeliveryRepository _repository;
        private readonly DataGridView _grid = CreateGrid();
        private readonly TextBox _txtFiltro = new() { Dock = DockStyle.Top, PlaceholderText = "Filtrar por nombre, apellido o código..." };
        private List<ClienteDelivery> _clientes = [];

        public SeleccionarClienteDeliveryDialog(IClienteDeliveryRepository repository)
        {
            _repository = repository;
            Text = "Seleccionar cliente delivery";
            Width = 860;
            Height = 520;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = SurfaceColor;

            var btnAceptar = CreateButton("Aceptar", AccentColor, (_, _) => Confirmar());
            var btnCancelar = CreateButton("Cancelar", Color.FromArgb(220, 53, 69), (_, _) => DialogResult = DialogResult.Cancel);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, Padding = new Padding(12), FlowDirection = FlowDirection.RightToLeft };
            actions.Controls.AddRange([btnCancelar, btnAceptar]);

            Controls.Add(_grid);
            Controls.Add(_txtFiltro);
            Controls.Add(actions);

            Load += async (_, _) => await CargarAsync();
            _txtFiltro.TextChanged += (_, _) => Filtrar();
            _grid.DoubleClick += (_, _) => Confirmar();
        }

        public ClienteDelivery? ClienteSeleccionado { get; private set; }

        private async Task CargarAsync()
        {
            _clientes = (await _repository.BuscarAsync(null, null, null, null)).ToList();
            Filtrar();
        }

        private void Filtrar()
        {
            var filtro = _txtFiltro.Text.Trim();
            var data = _clientes
                .Where(c =>
                    filtro.Length == 0 ||
                    c.CodigoDelivery.Contains(filtro, StringComparison.OrdinalIgnoreCase) ||
                    (c.Apellido ?? string.Empty).Contains(filtro, StringComparison.OrdinalIgnoreCase) ||
                    (c.Nombre ?? string.Empty).Contains(filtro, StringComparison.OrdinalIgnoreCase))
                .Select(c => new
                {
                    c.CodigoDelivery,
                    c.NumeroIdentidad,
                    Cliente = $"{c.Apellido} {c.Nombre}".Trim()
                })
                .ToList();

            _grid.DataSource = null;
            _grid.DataSource = data;
            if (_grid.Columns.Count > 0)
            {
                _grid.Columns[0].HeaderText = "Código";
                _grid.Columns[1].HeaderText = "Identidad";
                _grid.Columns[2].HeaderText = "Cliente";
                _grid.AutoResizeColumns();
            }
        }

        private void Confirmar()
        {
            var row = _grid.CurrentCell?.RowIndex ?? -1;
            if (row < 0 || row >= _grid.Rows.Count)
                return;

            var codigo = _grid.Rows[row].Cells[0].Value?.ToString();
            ClienteSeleccionado = _clientes.FirstOrDefault(c => string.Equals(c.CodigoDelivery, codigo, StringComparison.OrdinalIgnoreCase));
            if (ClienteSeleccionado is null)
                return;

            DialogResult = DialogResult.OK;
        }
    }
}
