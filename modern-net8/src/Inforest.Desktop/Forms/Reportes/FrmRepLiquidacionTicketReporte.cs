using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte de liquidación de cajero por ticketera.
/// Legacy: <c>frmRepLiquidacionTicket.frm</c>, <c>spRep_LiquidacionSuma</c>
/// Regla: BR-REP-015
/// </summary>
public sealed class FrmRepLiquidacionTicketReporte : Form
{
    private readonly ObtenerReporteLiquidacionTicketHandler _handler;
    private readonly ILogger<FrmRepLiquidacionTicketReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpFechaInicio = null!;
    private DateTimePicker _dtpHoraInicio = null!;
    private DateTimePicker _dtpFechaFin = null!;
    private DateTimePicker _dtpHoraFin = null!;
    private CheckBox _chkTodosTurnos = null!;
    private CheckBox _chkDiaContable = null!;
    private TextBox _txtTurno = null!;
    private TextBox _txtUsuario = null!;
    private Label _lblEstado = null!;

    public FrmRepLiquidacionTicketReporte(
        ObtenerReporteLiquidacionTicketHandler handler,
        ILogger<FrmRepLiquidacionTicketReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;

        InitializeComponent();
    }

    private async void BtnVer_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: true, imprimir: false, exportarCsv: false);
    private async void BtnImprimir_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: false, imprimir: true, exportarCsv: false);
    private async void BtnExportar_Click(object? sender, EventArgs e) => await EmitirReporteAsync(mostrar: false, imprimir: false, exportarCsv: true);

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir, bool exportarCsv)
    {
        var fechaInicio = _dtpFechaInicio.Value.Date.Add(_dtpHoraInicio.Value.TimeOfDay);
        var fechaFin = _dtpFechaFin.Value.Date.Add(_dtpHoraFin.Value.TimeOfDay);

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

            var parametros = new LiquidacionTicketParametros
            {
                TodosLosTurnos = _chkTodosTurnos.Checked,
                DiaContable = _chkDiaContable.Checked,
                Turno = _chkTodosTurnos.Checked ? string.Empty : _txtTurno.Text.Trim(),
                Usuario = _txtUsuario.Text.Trim(),
                FechaInicio = fechaInicio,
                FechaFin = fechaFin,
                SectorVenta = string.Empty
            };

            var resultado = await _handler.HandleAsync(new ObtenerReporteLiquidacionTicketQuery(parametros));
            if (resultado.Filas.Count == 0)
            {
                MessageBox.Show("No hay datos para mostrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                _lblEstado.Text = "Sin resultados.";
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
                viewer.AgregarDataSource("LiquidacionTicket", resultado.Filas);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} filas";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte de liquidación por ticketera");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private static void ExportarCsv(IReadOnlyList<LiquidacionTicketRow> filas)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"LiquidacionTicket_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        var lineas = new List<string>
        {
            "TipoPedido,Neto,Impuesto1,Impuesto2,Impuesto3,Venta,Documentos,TotalDocumentos"
        };

        lineas.AddRange(filas.Select(f => string.Join(",",
            Csv(f.TTipoPedido),
            f.NNeto.ToString("0.00"),
            f.NImpuesto1.ToString("0.00"),
            f.NImpuesto2.ToString("0.00"),
            f.NImpuesto3.ToString("0.00"),
            f.NVenta.ToString("0.00"),
            f.NTotalPromedio.ToString(),
            f.Total00.ToString())));

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
        Text = "Liquidación de Cajero por Ticketera";
        Size = new Size(700, 330);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 7,
            Padding = new Padding(12)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));

        _dtpFechaInicio = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _dtpHoraInicio = new DateTimePicker { Format = DateTimePickerFormat.Time, ShowUpDown = true, Value = DateTime.Today };
        _dtpFechaFin = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _dtpHoraFin = new DateTimePicker { Format = DateTimePickerFormat.Time, ShowUpDown = true, Value = DateTime.Today.AddHours(23).AddMinutes(59) };
        _chkTodosTurnos = new CheckBox { Text = "Todos los turnos", Checked = true, AutoSize = true };
        _chkDiaContable = new CheckBox { Text = "Usar día contable", AutoSize = true };
        _txtTurno = new TextBox { Enabled = false };
        _txtUsuario = new TextBox();
        _lblEstado = new Label { Dock = DockStyle.Fill, ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft };

        _chkTodosTurnos.CheckedChanged += (_, _) => ToggleTurno();

        var rangoInicio = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        rangoInicio.Controls.AddRange([_dtpFechaInicio, _dtpHoraInicio]);

        var rangoFin = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        rangoFin.Controls.AddRange([_dtpFechaFin, _dtpHoraFin]);

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

        panel.Controls.Add(new Label { Text = "Fecha/Hora inicio:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        panel.Controls.Add(rangoInicio, 1, 0);
        panel.SetColumnSpan(rangoInicio, 2);
        panel.Controls.Add(new Label { Text = "Fecha/Hora fin:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        panel.Controls.Add(rangoFin, 1, 1);
        panel.SetColumnSpan(rangoFin, 2);
        panel.Controls.Add(new Label { Text = "Turno:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        panel.Controls.Add(_txtTurno, 1, 2);
        panel.Controls.Add(_chkTodosTurnos, 2, 2);
        panel.Controls.Add(new Label { Text = "Usuario:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
        panel.Controls.Add(_txtUsuario, 1, 3);
        panel.Controls.Add(new Label { Text = "(opcional)", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 3);
        panel.Controls.Add(new Label { Text = "Opciones:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 4);
        panel.Controls.Add(_chkDiaContable, 1, 4);
        panel.Controls.Add(_lblEstado, 0, 5);
        panel.SetColumnSpan(_lblEstado, 2);
        panel.Controls.Add(acciones, 2, 6);

        Controls.Add(panel);
    }
}
