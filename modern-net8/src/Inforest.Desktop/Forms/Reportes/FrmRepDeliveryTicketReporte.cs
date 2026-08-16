using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte "Cierre de Cajeros Delivery" — pagos de pedidos delivery agrupados por caja/motorizado/tipo pago.
/// Legacy: <c>frmRepDeliveryTicket.frm</c> (InfoRest.vbp)
/// Regla: BR-REP-017
/// </summary>
public sealed class FrmRepDeliveryTicketReporte : Form
{
    private readonly ObtenerReporteDeliveryTicketHandler _handler;
    private readonly ILogger<FrmRepDeliveryTicketReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private CheckBox _chkTodosTurnos = null!;
    private TextBox _txtTurno = null!;
    private CheckBox _chkTodasCajas = null!;
    private TextBox _txtCaja = null!;
    private CheckBox _chkTodosMotorizados = null!;
    private TextBox _txtMotorizado = null!;
    private Label _lblEstado = null!;

    public FrmRepDeliveryTicketReporte(
        ObtenerReporteDeliveryTicketHandler handler,
        ILogger<FrmRepDeliveryTicketReporte> logger,
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
        // Validaciones equivalentes al Legacy cmdOpcion_Click
        if (!_chkTodosTurnos.Checked && string.IsNullOrWhiteSpace(_txtTurno.Text))
        {
            MessageBox.Show("Debe escoger un Turno.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var fechaInicio = _dtpInicio.Value;
        var fechaFin = _dtpFin.Value;
        if (_chkTodosTurnos.Checked && fechaInicio > fechaFin)
        {
            MessageBox.Show("Error en Rango de Fechas.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodasCajas.Checked && string.IsNullOrWhiteSpace(_txtCaja.Text))
        {
            MessageBox.Show("Debe escoger una Caja.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosMotorizados.Checked && string.IsNullOrWhiteSpace(_txtMotorizado.Text))
        {
            MessageBox.Show("Debe escoger un Motorizado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var parametros = new DeliveryTicketParametros
            {
                TodosTurnos = _chkTodosTurnos.Checked,
                Turno = _chkTodosTurnos.Checked ? string.Empty : _txtTurno.Text.Trim(),
                FechaInicio = fechaInicio,
                FechaFin = fechaFin,
                TodasLasCajas = _chkTodasCajas.Checked,
                Caja = _chkTodasCajas.Checked ? string.Empty : _txtCaja.Text.Trim(),
                TodosLosMotorizados = _chkTodosMotorizados.Checked,
                Motorizado = _chkTodosMotorizados.Checked ? string.Empty : _txtMotorizado.Text.Trim()
            };

            var resultado = await _handler.HandleAsync(new ObtenerReporteDeliveryTicketQuery(parametros));

            if (resultado.Filas.Count == 0)
            {
                _lblEstado.Text = "Sin resultados.";
                MessageBox.Show(
                    "No se tienen Comprobantes pendientes de cobro del Canal Delivery.",
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                viewer.AgregarDataSource("DeliveryTicket", resultado.Filas);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} filas";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte delivery ticket");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private static void ExportarCsv(IReadOnlyList<DeliveryTicketRow> filas)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"DeliveryTicket_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        var lineas = new List<string>
        {
            "Caja,TipoPago,Motorizado,Documento,FRegistro,NVenta,NMonto,NVuelto,Moneda,Tarjeta,Numero,OtroTipo"
        };
        lineas.AddRange(filas.Select(f => string.Join(",",
            Csv(f.TCaja),
            Csv(f.TipoPago),
            Csv(f.Motorizado),
            Csv(f.TDocumento),
            f.FRegistro.ToString("yyyy-MM-dd HH:mm"),
            f.NVenta.ToString("0.00"),
            f.NMonto.ToString("0.00"),
            f.NVuelto.ToString("0.00"),
            Csv(f.Mon),
            Csv(f.Tarjeta),
            Csv(f.TNumero),
            Csv(f.OtroTipo))));

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
        if (_chkTodosTurnos.Checked) _txtTurno.Text = string.Empty;
    }

    private void ToggleCaja()
    {
        _txtCaja.Enabled = !_chkTodasCajas.Checked;
        if (_chkTodasCajas.Checked) _txtCaja.Text = string.Empty;
    }

    private void ToggleMotorizado()
    {
        _txtMotorizado.Enabled = !_chkTodosMotorizados.Checked;
        if (_chkTodosMotorizados.Checked) _txtMotorizado.Text = string.Empty;
    }

    private void InitializeComponent()
    {
        Text = "Cierre de Cajeros Delivery";
        Size = new Size(560, 380);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 7,
            Padding = new Padding(12)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        _dtpInicio = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm", Value = DateTime.Today };
        _dtpFin = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm", Value = DateTime.Today.AddDays(1).AddSeconds(-1) };

        _chkTodosTurnos = new CheckBox { Text = "Todos los Turnos", Checked = true, AutoSize = true };
        _txtTurno = new TextBox { Enabled = false };
        _chkTodosTurnos.CheckedChanged += (_, _) => ToggleTurno();

        _chkTodasCajas = new CheckBox { Text = "Todas las Cajas", Checked = true, AutoSize = true };
        _txtCaja = new TextBox { Enabled = false };
        _chkTodasCajas.CheckedChanged += (_, _) => ToggleCaja();

        _chkTodosMotorizados = new CheckBox { Text = "Todos los Motorizados", Checked = true, AutoSize = true };
        _txtMotorizado = new TextBox { Enabled = false };
        _chkTodosMotorizados.CheckedChanged += (_, _) => ToggleMotorizado();

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

        panel.Controls.Add(_chkTodosTurnos, 0, 1);
        panel.Controls.Add(_txtTurno, 1, 1);

        panel.Controls.Add(_chkTodasCajas, 0, 2);
        panel.Controls.Add(_txtCaja, 1, 2);

        panel.Controls.Add(_chkTodosMotorizados, 0, 3);
        panel.Controls.Add(_txtMotorizado, 1, 3);

        panel.Controls.Add(_lblEstado, 0, 5);
        panel.SetColumnSpan(_lblEstado, 3);
        panel.Controls.Add(acciones, 3, 5);

        Controls.Add(panel);
    }
}
