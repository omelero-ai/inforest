using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte de control de transacciones (anulados, transferidos, facturados).
/// Legacy: <c>frmRepAnulado.frm</c> — "Control de Transacciones", <c>spRep_Anulacion</c>
/// Regla: BR-REP-014
/// </summary>
public sealed class FrmRepAnuladoReporte : Form
{
    private readonly ObtenerReporteAnulacionHandler _handler;
    private readonly ILogger<FrmRepAnuladoReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private CheckBox _chkFranjaHoraria = null!;
    private CheckBox _chkFacturados = null!;
    private CheckBox _chkAnulados = null!;
    private CheckBox _chkTransferidos = null!;
    private Label _lblEstado = null!;

    public FrmRepAnuladoReporte(
        ObtenerReporteAnulacionHandler handler,
        ILogger<FrmRepAnuladoReporte> logger,
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
        if (_dtpInicio.Value.Date > _dtpFin.Value.Date)
        {
            MessageBox.Show("La fecha inicio no puede ser mayor a la fecha fin.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkFacturados.Checked && !_chkAnulados.Checked && !_chkTransferidos.Checked)
        {
            MessageBox.Show("Debe seleccionar al menos un tipo de transacción.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var parametros = new AnulacionParametros
            {
                FranjaHoraria = _chkFranjaHoraria.Checked,
                Turno = string.Empty,
                FechaInicio = _dtpInicio.Value.Date,
                FechaFin = _dtpFin.Value.Date.AddDays(1).AddSeconds(-1),
                FlagFacturados = _chkFacturados.Checked,
                FlagAnulados = _chkAnulados.Checked,
                FlagTransferidos = _chkTransferidos.Checked,
                Criterio = string.Empty
            };

            var resultado = await _handler.HandleAsync(new ObtenerReporteAnulacionQuery(parametros));

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
                viewer.AgregarDataSource("Anulacion", resultado.Filas);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} registros";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte de anulaciones");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private static void ExportarCsv(IReadOnlyList<AnulacionRow> filas)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"ControlTransacciones_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        var lineas = new List<string>
        {
            "Pedido,Item,Producto,Cantidad,Venta,Estado,Documento,FechaRegistro,Imprime,MotivoAnulacion,Observacion,UsuarioAnulado,FechaAnulado,Turno,FechaItem"
        };

        lineas.AddRange(filas.Select(f => string.Join(",",
            Csv(f.TCodigoPedido),
            Csv(f.TItem),
            Csv(f.TCodigoProducto),
            f.NCantidad.ToString("0.##"),
            f.NVenta.ToString("0.00"),
            Csv(f.TEstadoItem),
            Csv(f.TDocumento),
            f.FRegistro.ToString("dd/MM/yyyy HH:mm"),
            f.LImprime ? "1" : "0",
            Csv(f.TMotivoAnulacion),
            Csv(f.TObservacionAnulado),
            Csv(f.TUsuarioAnulado),
            f.FRegAnulado.HasValue ? f.FRegAnulado.Value.ToString("dd/MM/yyyy HH:mm") : string.Empty,
            Csv(f.TTurno),
            f.FFechaItem.HasValue ? f.FFechaItem.Value.ToString("dd/MM/yyyy HH:mm") : string.Empty)));

        File.WriteAllLines(dialog.FileName, lineas);
    }

    private static string Csv(string valor)
        => $"\"{(valor ?? string.Empty).Replace("\"", "\"\"")}\"";

    private void SetBotones(bool enabled)
    {
        foreach (var boton in Controls.OfType<Button>())
            boton.Enabled = enabled;
    }

    private void InitializeComponent()
    {
        Text = "Control de Transacciones";
        Size = new Size(560, 320);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 6,
            Padding = new Padding(12)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _dtpInicio = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddMonths(-1) };
        _dtpFin = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _chkFranjaHoraria = new CheckBox { Text = "Franja Horaria", AutoSize = true };
        _chkFacturados = new CheckBox { Text = "Facturados", AutoSize = true, Checked = true };
        _chkAnulados = new CheckBox { Text = "Anulados", AutoSize = true, Checked = true };
        _chkTransferidos = new CheckBox { Text = "Transferidos", AutoSize = true, Checked = true };

        var grupoTipo = new GroupBox { Text = "Tipo de Transacción", Dock = DockStyle.Fill, Height = 75 };
        var flujoTipo = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
        flujoTipo.Controls.AddRange([_chkFacturados, _chkAnulados, _chkTransferidos]);
        grupoTipo.Controls.Add(flujoTipo);

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
        panel.Controls.Add(new Label { Text = "Fecha fin:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        panel.Controls.Add(_dtpFin, 1, 1);
        panel.Controls.Add(new Label { Text = "Opciones:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        panel.Controls.Add(_chkFranjaHoraria, 1, 2);
        panel.Controls.Add(grupoTipo, 0, 3);
        panel.SetColumnSpan(grupoTipo, 2);
        panel.Controls.Add(_lblEstado, 0, 4);
        panel.Controls.Add(acciones, 1, 5);

        Controls.Add(panel);
    }
}
