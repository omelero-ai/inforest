using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte paloteo de producción por ticketera.
/// Legacy: <c>frmRepPaloteoTicket.frm</c> (InfoRest.vbp)
/// Regla: BR-REP-016
/// </summary>
public sealed class FrmRepPaloteoTicketReporte : Form
{
    private readonly ObtenerReportePaloteoTicketHandler _handler;
    private readonly ILogger<FrmRepPaloteoTicketReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private CheckBox _chkTodosTurnos = null!;
    private TextBox _txtTurno = null!;
    private ComboBox _cmbOrigen = null!;
    private ComboBox _cmbOrden = null!;
    private CheckBox _chkSubtotales = null!;
    private Label _lblEstado = null!;

    private TextBox _txtSalon = null!;
    private TextBox _txtTipoProducto = null!;
    private TextBox _txtMozo = null!;
    private TextBox _txtTipoPedido = null!;
    private TextBox _txtOrigenVenta = null!;
    private TextBox _txtArea = null!;
    private TextBox _txtGrupo = null!;
    private TextBox _txtSubGrupo = null!;
    private TextBox _txtProducto = null!;
    private TextBox _txtCliente = null!;

    public FrmRepPaloteoTicketReporte(
        ObtenerReportePaloteoTicketHandler handler,
        ILogger<FrmRepPaloteoTicketReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        InitializeComponent();
    }

    private async void BtnEmitir_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: true, imprimir: false, exportarCsv: false);
    private async void BtnImprimir_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: false, imprimir: true, exportarCsv: false);
    private async void BtnExportar_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: false, imprimir: false, exportarCsv: true);

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir, bool exportarCsv)
    {
        var fechaInicio = _dtpInicio.Value.Date;
        var fechaFin = _dtpFin.Value.Date.AddDays(1).AddSeconds(-1);
        if (fechaInicio > fechaFin)
        {
            MessageBox.Show("La fecha inicio no puede ser mayor a la fecha fin.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosTurnos.Checked && string.IsNullOrWhiteSpace(_txtTurno.Text))
        {
            MessageBox.Show("Debe ingresar un turno cuando no selecciona todos los turnos.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var parametros = new PaloteoTicketParametros
            {
                TodosTurnos = _chkTodosTurnos.Checked,
                Turno = _chkTodosTurnos.Checked ? string.Empty : _txtTurno.Text.Trim(),
                FechaInicio = fechaInicio,
                FechaFin = fechaFin,
                Salon = _txtSalon.Text.Trim(),
                TipoProducto = _txtTipoProducto.Text.Trim(),
                Mozo = _txtMozo.Text.Trim(),
                TipoPedido = _txtTipoPedido.Text.Trim(),
                OrigenVenta = _txtOrigenVenta.Text.Trim(),
                Area = _txtArea.Text.Trim(),
                Grupo = _txtGrupo.Text.Trim(),
                SubGrupo = _txtSubGrupo.Text.Trim(),
                CodigoProducto = _txtProducto.Text.Trim(),
                CodigoCliente = _txtCliente.Text.Trim(),
                Origen = (OrigenPaloteoTicket)_cmbOrigen.SelectedValue,
                OrdenarPorCodigoProducto = string.Equals(_cmbOrden.SelectedValue?.ToString(), "CODIGO", StringComparison.Ordinal),
                MostrarTotalPorProducto = _chkSubtotales.Checked
            };

            var resultado = await _handler.HandleAsync(new ObtenerReportePaloteoTicketQuery(parametros));
            if (resultado.Filas.Count == 0)
            {
                _lblEstado.Text = "Sin resultados.";
                MessageBox.Show("No hay datos para mostrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (exportarCsv)
            {
                ExportarCsv(resultado.Filas);
            }
            else
            {
                using var viewer = new ReporteViewer(
                    Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                    _rutaPlantillas);
                viewer.CargarPlantilla(resultado.NombrePlantilla);
                viewer.AgregarDataSource("PaloteoTicket", resultado.Filas);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} filas";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte paloteo ticket");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private static void ExportarCsv(IReadOnlyList<PaloteoTicketRow> filas)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"PaloteoTicket_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        var lineas = new List<string>
        {
            "CodigoProducto,Local,Salon,TipoProducto,Grupo,SubGrupo,Producto,Cantidad,Venta"
        };
        lineas.AddRange(filas.Select(f => string.Join(",",
            Csv(f.TCodProducto),
            Csv(f.Local),
            Csv(f.Salon),
            Csv(f.TipoProducto),
            Csv(f.Grupo),
            Csv(f.SubGrupo),
            Csv(f.Producto),
            f.Cantidad.ToString("0.00"),
            f.Venta.ToString("0.00"))));

        File.WriteAllLines(dialog.FileName, lineas);
    }

    private static string Csv(string valor)
        => $"\"{(valor ?? string.Empty).Replace("\"", "\"\"")}\"";

    private void SetBotones(bool enabled)
    {
        foreach (var boton in Controls.OfType<Button>())
            boton.Enabled = enabled;
    }

    private void ToggleTurno()
    {
        _txtTurno.Enabled = !_chkTodosTurnos.Checked;
        if (_chkTodosTurnos.Checked)
            _txtTurno.Text = string.Empty;
    }

    private void InitializeComponent()
    {
        Text = "Paloteo de Producción por Ticketera";
        Size = new Size(860, 490);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 9,
            Padding = new Padding(12)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        _dtpInicio = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _dtpFin = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _chkTodosTurnos = new CheckBox { Text = "Todos los turnos", Checked = true, AutoSize = true };
        _txtTurno = new TextBox { Enabled = false };
        _chkTodosTurnos.CheckedChanged += (_, _) => ToggleTurno();

        _cmbOrigen = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        _cmbOrigen.DataSource = new[]
        {
            new { Codigo = OrigenPaloteoTicket.Produccion, Nombre = "Producción (a+b+c)" },
            new { Codigo = OrigenPaloteoTicket.Venta, Nombre = "Ventas" },
            new { Codigo = OrigenPaloteoTicket.Cortesia, Nombre = "Cortesías" },
            new { Codigo = OrigenPaloteoTicket.CuentaCorriente, Nombre = "Ctas. Corrientes" },
            new { Codigo = OrigenPaloteoTicket.Combinacion, Nombre = "Combinación" },
            new { Codigo = OrigenPaloteoTicket.Cargos, Nombre = "Cargos" },
            new { Codigo = OrigenPaloteoTicket.PedidosFacturados, Nombre = "Pedidos Facturados" }
        };
        _cmbOrigen.DisplayMember = "Nombre";
        _cmbOrigen.ValueMember = "Codigo";

        _cmbOrden = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        _cmbOrden.DataSource = new[]
        {
            new { Codigo = "ALFABETICO", Nombre = "Alfabético" },
            new { Codigo = "CODIGO", Nombre = "Código Producto" }
        };
        _cmbOrden.DisplayMember = "Nombre";
        _cmbOrden.ValueMember = "Codigo";

        _chkSubtotales = new CheckBox { Text = "Mostrar total por producto", AutoSize = true };

        _txtSalon = new TextBox();
        _txtTipoProducto = new TextBox();
        _txtMozo = new TextBox();
        _txtTipoPedido = new TextBox();
        _txtOrigenVenta = new TextBox();
        _txtArea = new TextBox();
        _txtGrupo = new TextBox();
        _txtSubGrupo = new TextBox();
        _txtProducto = new TextBox();
        _txtCliente = new TextBox();
        _lblEstado = new Label { Dock = DockStyle.Fill, ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft };

        var acciones = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft
        };
        var btnSalir = new Button { Text = "Salir", Width = 90 };
        var btnExportar = new Button { Text = "Exportar", Width = 90 };
        var btnImprimir = new Button { Text = "Imprimir", Width = 90 };
        var btnEmitir = new Button { Text = "Emitir", Width = 90 };
        btnSalir.Click += (_, _) => Close();
        btnExportar.Click += BtnExportar_Click;
        btnImprimir.Click += BtnImprimir_Click;
        btnEmitir.Click += BtnEmitir_Click;
        acciones.Controls.AddRange([btnSalir, btnExportar, btnImprimir, btnEmitir]);

        panel.Controls.Add(new Label { Text = "Fecha inicio:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        panel.Controls.Add(_dtpInicio, 1, 0);
        panel.Controls.Add(new Label { Text = "Fecha fin:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
        panel.Controls.Add(_dtpFin, 3, 0);

        panel.Controls.Add(new Label { Text = "Turno:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        panel.Controls.Add(_txtTurno, 1, 1);
        panel.Controls.Add(_chkTodosTurnos, 2, 1);

        panel.Controls.Add(new Label { Text = "Origen de datos:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        panel.Controls.Add(_cmbOrigen, 1, 2);
        panel.Controls.Add(new Label { Text = "Orden:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 2);
        panel.Controls.Add(_cmbOrden, 3, 2);

        panel.Controls.Add(new Label { Text = "Salón:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
        panel.Controls.Add(_txtSalon, 1, 3);
        panel.Controls.Add(new Label { Text = "Tipo producto:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 3);
        panel.Controls.Add(_txtTipoProducto, 3, 3);

        panel.Controls.Add(new Label { Text = "Mozo:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 4);
        panel.Controls.Add(_txtMozo, 1, 4);
        panel.Controls.Add(new Label { Text = "Tipo pedido:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 4);
        panel.Controls.Add(_txtTipoPedido, 3, 4);

        panel.Controls.Add(new Label { Text = "Origen venta:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 5);
        panel.Controls.Add(_txtOrigenVenta, 1, 5);
        panel.Controls.Add(new Label { Text = "Área:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 5);
        panel.Controls.Add(_txtArea, 3, 5);

        panel.Controls.Add(new Label { Text = "Grupo:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 6);
        panel.Controls.Add(_txtGrupo, 1, 6);
        panel.Controls.Add(new Label { Text = "Subgrupo:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 6);
        panel.Controls.Add(_txtSubGrupo, 3, 6);

        panel.Controls.Add(new Label { Text = "Producto:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 7);
        panel.Controls.Add(_txtProducto, 1, 7);
        panel.Controls.Add(new Label { Text = "Cliente:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 7);
        panel.Controls.Add(_txtCliente, 3, 7);

        panel.Controls.Add(_chkSubtotales, 0, 8);
        panel.Controls.Add(_lblEstado, 1, 8);
        panel.SetColumnSpan(_lblEstado, 2);
        panel.Controls.Add(acciones, 3, 8);

        Controls.Add(panel);
    }
}
