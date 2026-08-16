using Inforest.Application.Maestros;
using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Legacy: <c>frmRepCtaCte.frm</c>, <c>spRep_CtaCteN</c>, <c>vTipoCtaCte</c>, <c>vSubTipoCtaCte</c>
/// Regla: BR-REP-013
/// </summary>
public sealed class FrmRepCtaCteReporte : Form
{
    private readonly ObtenerReporteCtaCteOperativaHandler _handler;
    private readonly ObtenerCuentasCorrientesHandler _cuentasHandler;
    private readonly ObtenerTiposCtaCteReporteHandler _tiposHandler;
    private readonly ObtenerSubTiposCtaCteReporteHandler _subTiposHandler;
    private readonly ILogger<FrmRepCtaCteReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private ComboBox _cmbEstado = null!;
    private ComboBox _cmbCliente = null!;
    private ComboBox _cmbTipo = null!;
    private ComboBox _cmbSubTipo = null!;
    private CheckBox _chkEstado = null!;
    private CheckBox _chkCliente = null!;
    private CheckBox _chkTipo = null!;
    private CheckBox _chkSubTipo = null!;
    private RadioButton _optConsolidado = null!;
    private RadioButton _optResumido = null!;
    private RadioButton _optDetallado = null!;
    private Label _lblEstado = null!;

    public FrmRepCtaCteReporte(
        ObtenerReporteCtaCteOperativaHandler handler,
        ObtenerCuentasCorrientesHandler cuentasHandler,
        ObtenerTiposCtaCteReporteHandler tiposHandler,
        ObtenerSubTiposCtaCteReporteHandler subTiposHandler,
        ILogger<FrmRepCtaCteReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _cuentasHandler = cuentasHandler;
        _tiposHandler = tiposHandler;
        _subTiposHandler = subTiposHandler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;

        InitializeComponent();
        Load += async (_, _) => await CargarCatalogosAsync();
    }

    private async Task CargarCatalogosAsync()
    {
        try
        {
            _lblEstado.Text = "Cargando catálogos...";

            _cmbEstado.Items.Clear();
            _cmbEstado.Items.AddRange(["FACTURADO", "NO FACTURADO"]);
            _cmbEstado.SelectedIndex = 0;

            var cuentas = await _cuentasHandler.HandleAsync(new ObtenerCuentasCorrientesQuery());
            var clientes = cuentas.EsExitoso && cuentas.Valor is not null
                ? cuentas.Valor.Select(x => new ReporteFiltroOpcion
                {
                    Codigo = x.CodigoCtaCte,
                    Descripcion = $"{x.CodigoCtaCte} - {x.Descripcion}"
                }).ToList()
                : [];

            _cmbCliente.DataSource = clientes;
            _cmbCliente.DisplayMember = nameof(ReporteFiltroOpcion.Descripcion);
            _cmbCliente.ValueMember = nameof(ReporteFiltroOpcion.Codigo);

            var tipos = await _tiposHandler.HandleAsync();
            _cmbTipo.DataSource = tipos.ToList();
            _cmbTipo.DisplayMember = nameof(ReporteFiltroOpcion.Descripcion);
            _cmbTipo.ValueMember = nameof(ReporteFiltroOpcion.Codigo);

            await RecargarSubTiposAsync();
            _lblEstado.Text = "Listo.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error cargando catálogos de cuentas corrientes");
            _lblEstado.Text = "Error al cargar catálogos.";
            MessageBox.Show($"No se pudieron cargar los filtros: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async Task RecargarSubTiposAsync()
    {
        var tipo = _chkTipo.Checked ? string.Empty : (_cmbTipo.SelectedValue?.ToString() ?? string.Empty);
        var subTipos = await _subTiposHandler.HandleAsync(tipo);
        _cmbSubTipo.DataSource = subTipos.ToList();
        _cmbSubTipo.DisplayMember = nameof(ReporteFiltroOpcion.Descripcion);
        _cmbSubTipo.ValueMember = nameof(ReporteFiltroOpcion.Codigo);
    }

    private async void BtnVer_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: true, imprimir: false, exportarCsv: false);
    private async void BtnImprimir_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: false, imprimir: true, exportarCsv: false);
    private async void BtnExportar_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: false, imprimir: false, exportarCsv: true);

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir, bool exportarCsv)
    {
        if (_dtpInicio.Value.Date > _dtpFin.Value.Date)
        {
            MessageBox.Show("La fecha inicio no puede ser mayor a la fecha fin.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var parametros = new CtaCteOperativaParametros
            {
                FlagConsolidado = _optConsolidado.Checked,
                FlagResumido = _optResumido.Checked,
                FlagDetalle = _optDetallado.Checked,
                FechaInicio = _dtpInicio.Value.Date,
                FechaFin = _dtpFin.Value.Date.AddDays(1).AddSeconds(-1),
                Estado = _chkEstado.Checked ? string.Empty : (_cmbEstado.SelectedIndex == 0 ? "01" : "02"),
                Cliente = _chkCliente.Checked ? string.Empty : (_cmbCliente.SelectedValue?.ToString() ?? string.Empty),
                TipoCtaCte = _chkTipo.Checked ? string.Empty : (_cmbTipo.SelectedValue?.ToString() ?? string.Empty),
                SubTipoCtaCte = _chkSubTipo.Checked ? string.Empty : (_cmbSubTipo.SelectedValue?.ToString() ?? string.Empty)
            };

            var resultado = await _handler.HandleAsync(new ObtenerReporteCtaCteOperativaQuery(parametros));

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
                viewer.AgregarDataSource("CtaCte", resultado.Filas);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} registros";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte operativo de cuenta corriente");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private static void ExportarCsv(IReadOnlyList<CtaCteOperativaRow> filas)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"EstadosCtaCte_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        var lineas = new List<string>
        {
            "Cliente,Identidad,Consumo,Linea,Saldo,Pedido,Fecha,Local,Estado,Venta,Producto,Cantidad,Documento,Tipo,SubTipo"
        };

        lineas.AddRange(filas.Select(f => string.Join(",",
            Csv(f.Descripcion),
            Csv(f.Identidad),
            f.NConsumo != 0 ? f.NConsumo.ToString("0.00") : f.Consumo.ToString("0.00"),
            f.NLinea != 0 ? f.NLinea.ToString("0.00") : f.Linea.ToString("0.00"),
            f.NSaldo != 0 ? f.NSaldo.ToString("0.00") : f.Saldo.ToString("0.00"),
            Csv(f.TCodigoPedido),
            Csv((f.FFecha != default ? f.FFecha : f.Fecha).ToString("dd/MM/yyyy")),
            Csv(f.Local),
            Csv(f.TEstadoPedido),
            (f.NVenta != 0 ? f.NVenta : f.Suma).ToString("0.00"),
            Csv(f.Producto),
            f.NCantidad.ToString("0.##"),
            Csv(f.TDocumento),
            Csv(f.TTipoCtaCte),
            Csv(f.TSubTipoCtaCte))));

        File.WriteAllLines(dialog.FileName, lineas);
    }

    private static string Csv(string valor)
        => $"\"{(valor ?? string.Empty).Replace("\"", "\"\"")}\"";

    private void SetBotones(bool enabled)
    {
        foreach (var boton in Controls.OfType<Button>())
            boton.Enabled = enabled;
    }

    private void ToggleFiltro(CheckBox checkBox, Control control)
        => control.Enabled = !checkBox.Checked;

    private void InitializeComponent()
    {
        Text = "Estados de Cuentas Corrientes";
        Size = new Size(760, 360);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 8,
            Padding = new Padding(12)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));

        _dtpInicio = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddMonths(-1) };
        _dtpFin = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _cmbEstado = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
        _cmbCliente = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
        _cmbTipo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
        _cmbSubTipo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };

        _chkEstado = new CheckBox { Text = "Todos los estados", Checked = true, AutoSize = true };
        _chkCliente = new CheckBox { Text = "Todos los clientes", Checked = true, AutoSize = true };
        _chkTipo = new CheckBox { Text = "Todos los tipos", Checked = true, AutoSize = true };
        _chkSubTipo = new CheckBox { Text = "Todos los subtipos", Checked = true, AutoSize = true };

        _chkEstado.CheckedChanged += (_, _) => ToggleFiltro(_chkEstado, _cmbEstado);
        _chkCliente.CheckedChanged += (_, _) => ToggleFiltro(_chkCliente, _cmbCliente);
        _chkTipo.CheckedChanged += async (_, _) =>
        {
            ToggleFiltro(_chkTipo, _cmbTipo);
            await RecargarSubTiposAsync();
        };
        _chkSubTipo.CheckedChanged += (_, _) => ToggleFiltro(_chkSubTipo, _cmbSubTipo);
        _cmbTipo.SelectedValueChanged += async (_, _) =>
        {
            if (!_chkTipo.Checked)
                await RecargarSubTiposAsync();
        };

        var grupoReporte = new GroupBox { Text = "Tipo de reporte", Dock = DockStyle.Fill };
        var flujoReporte = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        _optConsolidado = new RadioButton { Text = "Consolidado", AutoSize = true };
        _optResumido = new RadioButton { Text = "Resumido", AutoSize = true };
        _optDetallado = new RadioButton { Text = "Detallado", AutoSize = true, Checked = true };
        flujoReporte.Controls.AddRange([_optConsolidado, _optResumido, _optDetallado]);
        grupoReporte.Controls.Add(flujoReporte);

        var acciones = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft
        };
        var btnSalir = new Button { Text = "Salir", Width = 90 };
        var btnExportar = new Button { Text = "Exportar", Width = 90 };
        var btnImprimir = new Button { Text = "Imprimir", Width = 90 };
        var btnVer = new Button { Text = "Emitir", Width = 90 };
        btnSalir.Click += (_, _) => Close();
        btnExportar.Click += BtnExportar_Click;
        btnImprimir.Click += BtnImprimir_Click;
        btnVer.Click += BtnVer_Click;
        acciones.Controls.AddRange([btnSalir, btnExportar, btnImprimir, btnVer]);

        _lblEstado = new Label { Dock = DockStyle.Fill, ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft };

        panel.Controls.Add(new Label { Text = "Fecha inicio:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        panel.Controls.Add(_dtpInicio, 1, 0);
        panel.Controls.Add(new Label(), 2, 0);
        panel.Controls.Add(new Label { Text = "Fecha fin:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        panel.Controls.Add(_dtpFin, 1, 1);
        panel.Controls.Add(new Label(), 2, 1);
        panel.Controls.Add(new Label { Text = "Estado:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        panel.Controls.Add(_cmbEstado, 1, 2);
        panel.Controls.Add(_chkEstado, 2, 2);
        panel.Controls.Add(new Label { Text = "Cliente:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
        panel.Controls.Add(_cmbCliente, 1, 3);
        panel.Controls.Add(_chkCliente, 2, 3);
        panel.Controls.Add(new Label { Text = "Tipo Cta Cte:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 4);
        panel.Controls.Add(_cmbTipo, 1, 4);
        panel.Controls.Add(_chkTipo, 2, 4);
        panel.Controls.Add(new Label { Text = "Subtipo Cta Cte:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 5);
        panel.Controls.Add(_cmbSubTipo, 1, 5);
        panel.Controls.Add(_chkSubTipo, 2, 5);
        panel.Controls.Add(grupoReporte, 1, 6);
        panel.SetColumnSpan(grupoReporte, 2);
        panel.Controls.Add(_lblEstado, 0, 7);
        panel.SetColumnSpan(_lblEstado, 2);
        panel.Controls.Add(acciones, 2, 7);

        Controls.Add(panel);
    }
}
