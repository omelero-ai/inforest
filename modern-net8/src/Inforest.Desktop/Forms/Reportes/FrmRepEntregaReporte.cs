using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte de Entregas (Delivery central pedidos).
/// Legacy: <c>frmRepEntrega.frm</c> (InfoRest.vbp)
/// Regla: BR-REP-019
/// </summary>
public sealed class FrmRepEntregaReporte : Form
{
    private readonly ObtenerReporteEntregaHandler _handler;
    private readonly ILogger<FrmRepEntregaReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private CheckBox _chkTodosCanales = null!;
    private TextBox _txtCanal = null!;
    private CheckBox _chkTodosClientes = null!;
    private TextBox _txtCliente = null!;
    private CheckBox _chkTodosGrupos = null!;
    private TextBox _txtGrupo = null!;
    private CheckBox _chkTodosSubGrupos = null!;
    private TextBox _txtSubGrupo = null!;
    private CheckBox _chkTodosProductos = null!;
    private TextBox _txtProducto = null!;
    private CheckBox _chkTodosEstadosEntrega = null!;
    private ComboBox _cboEstadoEntrega = null!;
    private CheckBox _chkTodosEstadosCancelacion = null!;
    private ComboBox _cboEstadoCancelacion = null!;
    private RadioButton _rdoFormato1 = null!;
    private RadioButton _rdoResumido = null!;
    private RadioButton _rdoFormato2 = null!;
    private Label _lblEstado = null!;

    public FrmRepEntregaReporte(
        ObtenerReporteEntregaHandler handler,
        ILogger<FrmRepEntregaReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        InitializeComponent();
    }

    private async void BtnEmitir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: true, imprimir: false, exportarCsv: false);

    private async void BtnImprimir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: true, exportarCsv: false);

    private async void BtnExportar_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: false, exportarCsv: true);

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir, bool exportarCsv)
    {
        if (_dtpInicio.Value > _dtpFin.Value)
        {
            MessageBox.Show("Error en Rango de Fechas.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosCanales.Checked && string.IsNullOrWhiteSpace(_txtCanal.Text))
        {
            MessageBox.Show("Debe escoger un Canal de Venta.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosClientes.Checked && string.IsNullOrWhiteSpace(_txtCliente.Text))
        {
            MessageBox.Show("Debe escoger un Cliente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosGrupos.Checked && string.IsNullOrWhiteSpace(_txtGrupo.Text))
        {
            MessageBox.Show("Debe escoger un Grupo.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosSubGrupos.Checked && string.IsNullOrWhiteSpace(_txtSubGrupo.Text))
        {
            MessageBox.Show("Debe escoger un SubGrupo.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosProductos.Checked && string.IsNullOrWhiteSpace(_txtProducto.Text))
        {
            MessageBox.Show("Debe escoger un Producto.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosEstadosEntrega.Checked && string.IsNullOrWhiteSpace(_cboEstadoEntrega.Text))
        {
            MessageBox.Show("Debe escoger un Estado de Entrega.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosEstadosCancelacion.Checked && string.IsNullOrWhiteSpace(_cboEstadoCancelacion.Text))
        {
            MessageBox.Show("Debe escoger un Estado de Cancelación.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var parametros = new EntregaParametros
            {
                FechaHoraInicio = _dtpInicio.Value,
                FechaHoraFin = _dtpFin.Value,
                CodigoCliente = _chkTodosClientes.Checked ? string.Empty : _txtCliente.Text.Trim(),
                EstadoEntrega = _chkTodosEstadosEntrega.Checked ? string.Empty : _cboEstadoEntrega.Text.Trim(),
                EstadoCancelacion = _chkTodosEstadosCancelacion.Checked ? string.Empty : _cboEstadoCancelacion.Text.Trim(),
                Grupo = _chkTodosGrupos.Checked ? string.Empty : _txtGrupo.Text.Trim(),
                SubGrupo = _chkTodosSubGrupos.Checked ? string.Empty : _txtSubGrupo.Text.Trim(),
                CodigoProducto = _chkTodosProductos.Checked ? string.Empty : _txtProducto.Text.Trim(),
                CanalVenta = _chkTodosCanales.Checked ? string.Empty : _txtCanal.Text.Trim(),
                Formato = ObtenerFormato()
            };

            var resultado = await _handler.HandleAsync(new ObtenerReporteEntregaQuery(parametros));

            if (resultado.Filas.Count == 0)
            {
                _lblEstado.Text = "Sin resultados.";
                MessageBox.Show("No hay Datos para Mostrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (exportarCsv)
            {
                ExportarCsv(resultado.Filas, parametros.Formato);
            }
            else
            {
                using var viewer = new ReporteViewer(
                    Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                    _rutaPlantillas);
                viewer.CargarPlantilla(resultado.NombrePlantilla);
                viewer.AgregarDataSource("Entregas", resultado.Filas);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} filas";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte de entregas");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private FormatoReporteEntrega ObtenerFormato()
    {
        if (_rdoFormato1.Checked) return FormatoReporteEntrega.DetalladoFormato1;
        if (_rdoFormato2.Checked) return FormatoReporteEntrega.DetalladoFormato2;
        return FormatoReporteEntrega.ResumidoPorProducto;
    }

    private static void ExportarCsv(IReadOnlyList<EntregaRow> filas, FormatoReporteEntrega formato)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"Entregas_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        var lineas = new List<string>();
        if (formato == FormatoReporteEntrega.ResumidoPorProducto)
        {
            lineas.Add("Producto,Cantidad");
            lineas.AddRange(filas.Select(f => string.Join(",",
                Csv(f.Producto),
                f.Cantidad.ToString("0.00"))));
        }
        else
        {
            lineas.Add("Pedido,FechaPedido,FechaEntrega,HoraEntrega,Cliente,Producto,Cantidad,Monto,SaldoPendiente,EstadoPedido,Cancelacion,CanalVenta");
            lineas.AddRange(filas.Select(f => string.Join(",",
                Csv(f.Pedido),
                Csv(f.FechaPedido),
                Csv(f.FechaEntrega),
                Csv(f.HoraEntrega),
                Csv(f.Cliente),
                Csv(f.Producto),
                f.Cantidad.ToString("0.00"),
                f.Monto.ToString("0.00"),
                f.SaldoPendiente.ToString("0.00"),
                Csv(f.EstadoPedido),
                Csv(f.Cancelacion),
                Csv(f.CanalVenta))));
        }

        File.WriteAllLines(dialog.FileName, lineas);
    }

    private static string Csv(string valor)
        => $"\"{(valor ?? string.Empty).Replace("\"", "\"\"")}\"";

    private void SetBotones(bool enabled)
    {
        foreach (var boton in ObtenerBotones(this))
            boton.Enabled = enabled;
    }

    private static IEnumerable<Button> ObtenerBotones(Control parent)
    {
        foreach (Control control in parent.Controls)
        {
            if (control is Button button)
                yield return button;

            foreach (var nested in ObtenerBotones(control))
                yield return nested;
        }
    }

    private void ToggleControl(CheckBox checkBox, Control control)
    {
        control.Enabled = !checkBox.Checked;
        if (checkBox.Checked)
        {
            switch (control)
            {
                case TextBox textBox:
                    textBox.Text = string.Empty;
                    break;
                case ComboBox comboBox:
                    comboBox.SelectedIndex = -1;
                    comboBox.Text = string.Empty;
                    break;
            }
        }
    }

    private void InitializeComponent()
    {
        Text = "Reporte de Entregas";
        Size = new Size(860, 560);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 12,
            Padding = new Padding(12)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        _dtpInicio = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm", Value = DateTime.Today };
        _dtpFin = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm", Value = DateTime.Today.AddDays(1).AddSeconds(-1) };

        _chkTodosCanales = new CheckBox { Text = "Todos los Canales", Checked = true, AutoSize = true };
        _txtCanal = new TextBox { Enabled = false };
        _chkTodosCanales.CheckedChanged += (_, _) => ToggleControl(_chkTodosCanales, _txtCanal);

        _chkTodosClientes = new CheckBox { Text = "Todos los Clientes", Checked = true, AutoSize = true };
        _txtCliente = new TextBox { Enabled = false };
        _chkTodosClientes.CheckedChanged += (_, _) => ToggleControl(_chkTodosClientes, _txtCliente);

        _chkTodosGrupos = new CheckBox { Text = "Todos los Grupos", Checked = true, AutoSize = true };
        _txtGrupo = new TextBox { Enabled = false };
        _chkTodosGrupos.CheckedChanged += (_, _) => ToggleControl(_chkTodosGrupos, _txtGrupo);

        _chkTodosSubGrupos = new CheckBox { Text = "Todos los SubGrupos", Checked = true, AutoSize = true };
        _txtSubGrupo = new TextBox { Enabled = false };
        _chkTodosSubGrupos.CheckedChanged += (_, _) => ToggleControl(_chkTodosSubGrupos, _txtSubGrupo);

        _chkTodosProductos = new CheckBox { Text = "Todos los Productos", Checked = true, AutoSize = true };
        _txtProducto = new TextBox { Enabled = false };
        _chkTodosProductos.CheckedChanged += (_, _) => ToggleControl(_chkTodosProductos, _txtProducto);

        _chkTodosEstadosEntrega = new CheckBox { Text = "Todos los Estados de Entrega", Checked = true, AutoSize = true };
        _cboEstadoEntrega = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
        _cboEstadoEntrega.Items.AddRange(["ENTREGADO", "PENDIENTE"]);
        _chkTodosEstadosEntrega.CheckedChanged += (_, _) => ToggleControl(_chkTodosEstadosEntrega, _cboEstadoEntrega);

        _chkTodosEstadosCancelacion = new CheckBox { Text = "Todos los Estados de Cancelación", Checked = true, AutoSize = true };
        _cboEstadoCancelacion = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
        _cboEstadoCancelacion.Items.AddRange(["ANTICIPO", "NO PAGADO", "PAGADO"]);
        _chkTodosEstadosCancelacion.CheckedChanged += (_, _) => ToggleControl(_chkTodosEstadosCancelacion, _cboEstadoCancelacion);

        _rdoFormato1 = new RadioButton { Text = "Detallado por Pedido - Formato 1", AutoSize = true };
        _rdoResumido = new RadioButton { Text = "Resumido por Producto", AutoSize = true, Checked = true };
        _rdoFormato2 = new RadioButton { Text = "Formato 2", AutoSize = true };

        var panelFormato = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight
        };
        panelFormato.Controls.AddRange([_rdoFormato1, _rdoResumido, _rdoFormato2]);

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

        panel.Controls.Add(new Label { Text = "Fecha/hora inicio:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        panel.Controls.Add(_dtpInicio, 1, 0);
        panel.Controls.Add(new Label { Text = "Fecha/hora fin:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
        panel.Controls.Add(_dtpFin, 3, 0);

        panel.Controls.Add(_chkTodosCanales, 0, 1);
        panel.Controls.Add(_txtCanal, 1, 1);

        panel.Controls.Add(_chkTodosClientes, 2, 1);
        panel.Controls.Add(_txtCliente, 3, 1);

        panel.Controls.Add(_chkTodosGrupos, 0, 2);
        panel.Controls.Add(_txtGrupo, 1, 2);

        panel.Controls.Add(_chkTodosSubGrupos, 2, 2);
        panel.Controls.Add(_txtSubGrupo, 3, 2);

        panel.Controls.Add(_chkTodosProductos, 0, 3);
        panel.Controls.Add(_txtProducto, 1, 3);

        panel.Controls.Add(_chkTodosEstadosEntrega, 2, 3);
        panel.Controls.Add(_cboEstadoEntrega, 3, 3);

        panel.Controls.Add(_chkTodosEstadosCancelacion, 0, 4);
        panel.Controls.Add(_cboEstadoCancelacion, 1, 4);

        panel.Controls.Add(new Label { Text = "Formato:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 5);
        panel.Controls.Add(panelFormato, 1, 5);
        panel.SetColumnSpan(panelFormato, 3);

        panel.Controls.Add(_lblEstado, 0, 10);
        panel.SetColumnSpan(_lblEstado, 3);
        panel.Controls.Add(acciones, 3, 10);

        Controls.Add(panel);
    }
}
