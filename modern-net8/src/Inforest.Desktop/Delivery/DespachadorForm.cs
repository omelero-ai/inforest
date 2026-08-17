using System.Text;
using Inforest.Application.Despacho;
using Inforest.Application.Interfaces;
using Inforest.Desktop.Shared;
using Inforest.Domain.Entities.Delivery;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario operativo de despacho delivery.
/// Legacy: <c>frmDespachador.frm</c>.
/// Reglas: BR-DEL-DESP-001, BR-DEL-DESP-002, BR-DEL-DESP-003, BR-DEL-DESP-004.
/// </summary>
public sealed class DespachadorForm : Form
{
    private readonly ObtenerPedidosDespachadorHandler _obtenerHandler;
    private readonly ExportarDespachadorHandler _exportarHandler;
    private readonly ObtenerMotorizadosDespachoHandler _motorizadosHandler;
    private readonly ObtenerEmpacadoresDespachoHandler _empacadoresHandler;
    private readonly AsignarMotorizadoDespachoHandler _asignarMotorizadoHandler;
    private readonly DesasignarMotorizadoDespachoHandler _desasignarMotorizadoHandler;
    private readonly AsignarEmpacadorDespachoHandler _asignarEmpacadorHandler;
    private readonly DesasignarEmpacadorDespachoHandler _desasignarEmpacadorHandler;
    private readonly ISessionService _sessionService;

    private readonly BindingSource _bindingSource = new();
    private IReadOnlyList<PedidoDespachadorResumen> _pedidos = [];

    private DataGridView _grid = null!;
    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private Label _lblRegistro = null!;
    private System.Windows.Forms.Timer _timer = null!;

    public DespachadorForm(
        ObtenerPedidosDespachadorHandler obtenerHandler,
        ExportarDespachadorHandler exportarHandler,
        ObtenerMotorizadosDespachoHandler motorizadosHandler,
        ObtenerEmpacadoresDespachoHandler empacadoresHandler,
        AsignarMotorizadoDespachoHandler asignarMotorizadoHandler,
        DesasignarMotorizadoDespachoHandler desasignarMotorizadoHandler,
        AsignarEmpacadorDespachoHandler asignarEmpacadorHandler,
        DesasignarEmpacadorDespachoHandler desasignarEmpacadorHandler,
        ISessionService sessionService)
    {
        _obtenerHandler = obtenerHandler;
        _exportarHandler = exportarHandler;
        _motorizadosHandler = motorizadosHandler;
        _empacadoresHandler = empacadoresHandler;
        _asignarMotorizadoHandler = asignarMotorizadoHandler;
        _desasignarMotorizadoHandler = desasignarMotorizadoHandler;
        _asignarEmpacadorHandler = asignarEmpacadorHandler;
        _desasignarEmpacadorHandler = desasignarEmpacadorHandler;
        _sessionService = sessionService;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Despachador";
        WindowState = FormWindowState.Maximized;
        MinimumSize = new Size(1200, 650);
        BackColor = Color.FromArgb(245, 246, 247);

        var panelFiltros = new Panel
        {
            Dock = DockStyle.Top,
            Height = 56,
            Padding = new Padding(16, 12, 16, 8),
            BackColor = Color.White
        };

        panelFiltros.Controls.Add(new Label { Text = "Del", Left = 8, Top = 18, Width = 30 });
        _dtpInicio = new DateTimePicker { Left = 42, Top = 14, Width = 130, Value = DateTime.Today };
        panelFiltros.Controls.Add(_dtpInicio);

        panelFiltros.Controls.Add(new Label { Text = "al", Left = 178, Top = 18, Width = 20 });
        _dtpFin = new DateTimePicker { Left = 202, Top = 14, Width = 130, Value = DateTime.Today };
        panelFiltros.Controls.Add(_dtpFin);

        var btnFiltrar = CrearBoton("Filtrar", 350, Color.FromArgb(14, 116, 144));
        var btnExportar = CrearBoton("Exportar HTML", 450, Color.FromArgb(107, 114, 128), 120);
        var btnVista = CrearBoton("Vista previa", 580, Color.FromArgb(22, 163, 74), 110);
        var btnDetalle = CrearBoton("Detalle", 700, Color.FromArgb(59, 130, 246));
        var btnSalir = CrearBoton("Salir", 790, Color.FromArgb(220, 53, 69));
        panelFiltros.Controls.AddRange([btnFiltrar, btnExportar, btnVista, btnDetalle, btnSalir]);

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            AutoGenerateColumns = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None
        };
        _grid.Columns.AddRange(
            new DataGridViewTextBoxColumn { HeaderText = "Pedido", DataPropertyName = nameof(PedidoDespachadorResumen.CodigoPedido), Width = 90 },
            new DataGridViewTextBoxColumn { HeaderText = "Fecha", DataPropertyName = nameof(PedidoDespachadorResumen.FechaRegistro), Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM HH:mm" } },
            new DataGridViewTextBoxColumn { HeaderText = "Usuario", DataPropertyName = nameof(PedidoDespachadorResumen.Usuario), Width = 90 },
            new DataGridViewTextBoxColumn { HeaderText = "Caja", DataPropertyName = nameof(PedidoDespachadorResumen.Caja), Width = 70 },
            new DataGridViewTextBoxColumn { HeaderText = "Cliente", DataPropertyName = nameof(PedidoDespachadorResumen.Cliente), Width = 180 },
            new DataGridViewTextBoxColumn { HeaderText = "Teléfono", DataPropertyName = nameof(PedidoDespachadorResumen.Telefono), Width = 110 },
            new DataGridViewTextBoxColumn { HeaderText = "Dirección", DataPropertyName = nameof(PedidoDespachadorResumen.Direccion), Width = 220 },
            new DataGridViewTextBoxColumn { HeaderText = "Referencia", DataPropertyName = nameof(PedidoDespachadorResumen.Referencia), Width = 130 },
            new DataGridViewTextBoxColumn { HeaderText = "Empacador", DataPropertyName = nameof(PedidoDespachadorResumen.Empacador), Width = 110 },
            new DataGridViewTextBoxColumn { HeaderText = "Motorizado", DataPropertyName = nameof(PedidoDespachadorResumen.Motorizado), Width = 110 },
            new DataGridViewTextBoxColumn { HeaderText = "H.Asigna", DataPropertyName = nameof(PedidoDespachadorResumen.FechaAsignacion), Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM HH:mm" } },
            new DataGridViewTextBoxColumn { HeaderText = "Monto", DataPropertyName = nameof(PedidoDespachadorResumen.MontoTotal), Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } });

        var panelAcciones = new Panel
        {
            Dock = DockStyle.Right,
            Width = 165,
            Padding = new Padding(8),
            BackColor = Color.White
        };
        var btnAsignarMot = CrearBoton("Mot", 10, Color.FromArgb(14, 116, 144), 140, 0, 38);
        var btnQuitarMot = CrearBoton("No Mot", 10, Color.FromArgb(107, 114, 128), 140, 45, 38);
        var btnAsignarEmp = CrearBoton("Emp", 10, Color.FromArgb(22, 163, 74), 140, 90, 38);
        var btnQuitarEmp = CrearBoton("No Emp", 10, Color.FromArgb(107, 114, 128), 140, 135, 38);
        panelAcciones.Controls.AddRange([btnAsignarMot, btnQuitarMot, btnAsignarEmp, btnQuitarEmp]);

        var panelPie = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 48,
            Padding = new Padding(12, 8, 12, 8),
            BackColor = Color.White
        };
        var btnPrimero = new Button { Text = "|<", Left = 10, Top = 8, Width = 36, Height = 30 };
        var btnAnterior = new Button { Text = "<", Left = 50, Top = 8, Width = 36, Height = 30 };
        var btnSiguiente = new Button { Text = ">", Left = 90, Top = 8, Width = 36, Height = 30 };
        var btnUltimo = new Button { Text = ">|", Left = 130, Top = 8, Width = 36, Height = 30 };
        _lblRegistro = new Label { Left = 178, Top = 14, Width = 260, Text = "Registro 0 de 0", Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
        panelPie.Controls.AddRange([btnPrimero, btnAnterior, btnSiguiente, btnUltimo, _lblRegistro]);

        Controls.Add(_grid);
        Controls.Add(panelAcciones);
        Controls.Add(panelPie);
        Controls.Add(panelFiltros);

        Load += async (_, _) => await RecargarAsync();
        _grid.SelectionChanged += (_, _) => ActualizarRegistro();
        btnFiltrar.Click += async (_, _) => await RecargarAsync();
        btnExportar.Click += async (_, _) => await ExportarAsync();
        btnVista.Click += (_, _) => MostrarVistaPrevia();
        btnDetalle.Click += (_, _) => MostrarDetalle();
        btnSalir.Click += (_, _) => Close();
        btnAsignarMot.Click += async (_, _) => await AsignarMotorizadoAsync();
        btnQuitarMot.Click += async (_, _) => await DesasignarMotorizadoAsync();
        btnAsignarEmp.Click += async (_, _) => await AsignarEmpacadorAsync();
        btnQuitarEmp.Click += async (_, _) => await DesasignarEmpacadorAsync();
        btnPrimero.Click += (_, _) => SeleccionarIndice(0);
        btnAnterior.Click += (_, _) => SeleccionarIndice(_grid.CurrentRow is null ? 0 : Math.Max(0, _grid.CurrentRow.Index - 1));
        btnSiguiente.Click += (_, _) => SeleccionarIndice(_grid.CurrentRow is null ? 0 : Math.Min(_grid.Rows.Count - 1, _grid.CurrentRow.Index + 1));
        btnUltimo.Click += (_, _) => SeleccionarIndice(_grid.Rows.Count - 1);
        _grid.CellDoubleClick += (_, _) => MostrarDetalle();

        _timer = new System.Windows.Forms.Timer { Interval = 30000 };
        _timer.Tick += async (_, _) =>
        {
            if (!Visible) return;
            var pedidoActual = ObtenerPedidoSeleccionado()?.CodigoPedido;
            await RecargarAsync();
            if (!string.IsNullOrWhiteSpace(pedidoActual))
                SeleccionarPedido(pedidoActual);
        };
        _timer.Start();
    }

    private static Button CrearBoton(string texto, int left, Color color, int width = 90, int top = 11, int height = 32)
    {
        var button = new Button
        {
            Text = texto,
            Left = left,
            Top = top,
            Width = width,
            Height = height,
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        button.FlatAppearance.BorderSize = 0;
        return button;
    }

    private async Task RecargarAsync()
    {
        Cursor = Cursors.WaitCursor;
        try
        {
            var inicio = _dtpInicio.Value.Date;
            var fin = _dtpFin.Value.Date.AddHours(23).AddMinutes(59);
            var result = await _obtenerHandler.HandleAsync(new ObtenerPedidosDespachadorQuery(inicio, fin));
            if (!result.EsExitoso || result.Valor is null)
            {
                MessageBox.Show(result.MensajeError ?? "No se pudo cargar el panel de despacho.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _pedidos = result.Valor.ToList();
            _bindingSource.DataSource = _pedidos.ToList();
            _grid.DataSource = _bindingSource;
            SeleccionarIndice(0);
            ActualizarRegistro();
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private async Task AsignarMotorizadoAsync()
    {
        var pedido = ObtenerPedidoSeleccionado();
        if (pedido is null)
            return;

        var motorizados = await _motorizadosHandler.HandleAsync(new ObtenerMotorizadosDespachoQuery());
        if (!motorizados.EsExitoso || motorizados.Valor is null || motorizados.Valor.Count == 0)
        {
            MessageBox.Show(motorizados.MensajeError ?? "No existen motorizados activos.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var dlg = new FrmBusquedaRapida(
            "Seleccionar motorizado",
            motorizados.Valor.Select(x => new BusquedaItem(x.Codigo, x.Descripcion)).ToList(),
            "Código",
            "Motorizado");

        if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Resultado is null)
            return;

        var usuario = _sessionService.SesionActual?.CodigoUsuario ?? "SISTEMA";
        var result = await _asignarMotorizadoHandler.HandleAsync(
            new AsignarMotorizadoDespachoCommand(pedido.CodigoPedido, dlg.Resultado.Codigo, usuario));

        if (!result.EsExitoso && result.CodigoError == "DESPACHADOR_SUPERA_MONTO_MAXIMO")
        {
            var confirma = MessageBox.Show(
                "Con la asignación de este pedido se supera el monto máximo de asignación por motorizado. ¿Desea continuar?",
                Text,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirma == DialogResult.Yes)
            {
                result = await _asignarMotorizadoHandler.HandleAsync(
                    new AsignarMotorizadoDespachoCommand(pedido.CodigoPedido, dlg.Resultado.Codigo, usuario, AutorizarTarifaExtra: true));
            }
        }

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        await RecargarAsync();
        SeleccionarPedido(pedido.CodigoPedido);
    }

    private async Task DesasignarMotorizadoAsync()
    {
        var pedido = ObtenerPedidoSeleccionado();
        if (pedido is null)
            return;

        if (MessageBox.Show("¿Seguro de desasignar el motorizado?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        var result = await _desasignarMotorizadoHandler.HandleAsync(new DesasignarMotorizadoDespachoCommand(pedido.CodigoPedido));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        await RecargarAsync();
        SeleccionarPedido(pedido.CodigoPedido);
    }

    private async Task AsignarEmpacadorAsync()
    {
        var pedido = ObtenerPedidoSeleccionado();
        if (pedido is null)
            return;

        var empacadores = await _empacadoresHandler.HandleAsync(new ObtenerEmpacadoresDespachoQuery());
        if (!empacadores.EsExitoso || empacadores.Valor is null || empacadores.Valor.Count == 0)
        {
            MessageBox.Show(empacadores.MensajeError ?? "No existen empacadores activos.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var dlg = new FrmBusquedaRapida(
            "Seleccionar empacador",
            empacadores.Valor.Select(x => new BusquedaItem(x.Codigo, x.Descripcion)).ToList(),
            "Código",
            "Descripción");

        if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Resultado is null)
            return;

        var result = await _asignarEmpacadorHandler.HandleAsync(
            new AsignarEmpacadorDespachoCommand(pedido.CodigoPedido, dlg.Resultado.Codigo));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        await RecargarAsync();
        SeleccionarPedido(pedido.CodigoPedido);
    }

    private async Task DesasignarEmpacadorAsync()
    {
        var pedido = ObtenerPedidoSeleccionado();
        if (pedido is null)
            return;

        if (MessageBox.Show("¿Seguro de desasignar el empacador?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        var result = await _desasignarEmpacadorHandler.HandleAsync(new DesasignarEmpacadorDespachoCommand(pedido.CodigoPedido));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        await RecargarAsync();
        SeleccionarPedido(pedido.CodigoPedido);
    }

    private void MostrarVistaPrevia()
    {
        using var preview = new Form
        {
            Text = "Vista previa de despacho",
            Size = new Size(1100, 620),
            StartPosition = FormStartPosition.CenterParent
        };

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoGenerateColumns = false,
            DataSource = _pedidos.ToList()
        };
        foreach (DataGridViewColumn column in _grid.Columns)
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = column.HeaderText,
                DataPropertyName = column.DataPropertyName,
                Width = column.Width
            });
        }
        preview.Controls.Add(grid);
        preview.ShowDialog(this);
    }

    private void MostrarDetalle()
    {
        var pedido = ObtenerPedidoSeleccionado();
        if (pedido is null)
            return;

        var detalle = new StringBuilder();
        detalle.AppendLine($"Pedido: {pedido.CodigoPedido}");
        detalle.AppendLine($"Fecha: {pedido.FechaRegistro:dd/MM/yyyy HH:mm}");
        detalle.AppendLine($"Cliente: {pedido.Cliente}");
        detalle.AppendLine($"Teléfono: {pedido.Telefono}");
        detalle.AppendLine($"Dirección: {pedido.Direccion}");
        detalle.AppendLine($"Referencia: {pedido.Referencia}");
        detalle.AppendLine($"Empacador: {pedido.Empacador}");
        detalle.AppendLine($"Motorizado: {pedido.Motorizado}");
        detalle.AppendLine($"H.Asigna: {(pedido.FechaAsignacion.HasValue ? pedido.FechaAsignacion.Value.ToString("dd/MM/yyyy HH:mm") : "-")}");
        detalle.AppendLine($"Monto: {pedido.MontoTotal:N2}");

        MessageBox.Show(detalle.ToString(), "Detalle del pedido", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private async Task ExportarAsync()
    {
        var result = await _exportarHandler.HandleAsync(
            new ExportarDespachadorQuery(_dtpInicio.Value.Date, _dtpFin.Value.Date.AddHours(23).AddMinutes(59)));

        if (!result.EsExitoso || result.Valor is null)
        {
            MessageBox.Show(result.MensajeError ?? "No se pudo exportar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Filter = "Hyper Text Markup Language|*.html",
            FileName = $"Despachador_{DateTime.Now:yyyyMMdd_HHmmss}.html"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        var html = GenerarHtml(result.Valor);
        await File.WriteAllTextAsync(dialog.FileName, html, Encoding.UTF8);
        MessageBox.Show("Exportación completada.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static string GenerarHtml(IEnumerable<PedidoDespachadorResumen> pedidos)
    {
        var rows = string.Join(
            Environment.NewLine,
            pedidos.Select(p => $"<tr><td>{p.CodigoPedido}</td><td>{p.FechaRegistro:dd/MM/yyyy HH:mm}</td><td>{p.Cliente}</td><td>{p.Telefono}</td><td>{p.Empacador}</td><td>{p.Motorizado}</td><td>{p.MontoTotal:N2}</td></tr>"));
        return """
            <html>
            <head><meta charset="utf-8" /><title>Despachador</title></head>
            <body>
            <table border='1' cellspacing='0' cellpadding='4'>
                <tr>
                    <th>Pedido</th><th>Fecha</th><th>Cliente</th><th>Teléfono</th><th>Empacador</th><th>Motorizado</th><th>Monto</th>
                </tr>
            """ + rows + """
            </table>
            </body>
            </html>
            """;
    }

    private PedidoDespachadorResumen? ObtenerPedidoSeleccionado()
        => _grid.CurrentRow?.DataBoundItem as PedidoDespachadorResumen;

    private void SeleccionarIndice(int index)
    {
        if (_grid.Rows.Count == 0 || index < 0 || index >= _grid.Rows.Count)
            return;

        _grid.ClearSelection();
        _grid.Rows[index].Selected = true;
        _grid.CurrentCell = _grid.Rows[index].Cells[0];
        ActualizarRegistro();
    }

    private void SeleccionarPedido(string codigoPedido)
    {
        for (var i = 0; i < _grid.Rows.Count; i++)
        {
            if (_grid.Rows[i].DataBoundItem is PedidoDespachadorResumen item &&
                string.Equals(item.CodigoPedido, codigoPedido, StringComparison.OrdinalIgnoreCase))
            {
                SeleccionarIndice(i);
                return;
            }
        }
    }

    private void ActualizarRegistro()
    {
        var total = _grid.Rows.Count;
        var actual = _grid.CurrentRow is null ? 0 : _grid.CurrentRow.Index + 1;
        _lblRegistro.Text = $"Registro {actual} de {total}";
    }
}
