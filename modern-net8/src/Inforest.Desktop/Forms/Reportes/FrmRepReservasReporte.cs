using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte de Reservas — filtrado por rango de fecha/hora, estado y criterio de orden.
/// Legacy: <c>frmRepReservas.frm</c> (InfoRest.vbp)
/// Regla: BR-REP-018
/// </summary>
public sealed class FrmRepReservasReporte : Form
{
    private readonly ObtenerReporteReservasHandler _handler;
    private readonly ILogger<FrmRepReservasReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpFechaInicio = null!;
    private DateTimePicker _dtpHoraInicio = null!;
    private DateTimePicker _dtpFechaFin = null!;
    private DateTimePicker _dtpHoraFin = null!;
    private CheckBox _chkGenerado = null!;
    private CheckBox _chkAtendido = null!;
    private CheckBox _chkAnulado = null!;
    private ComboBox _cboOrden = null!;
    private Label _lblEstado = null!;

    public FrmRepReservasReporte(
        ObtenerReporteReservasHandler handler,
        ILogger<FrmRepReservasReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        InitializeComponent();
    }

    private async void BtnEmitir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync();

    private async Task EmitirReporteAsync()
    {
        // Validaciones equivalentes al Legacy cmdOpcion_Click (Case 0)
        var fechaHoraInicio = _dtpFechaInicio.Value.Date.Add(_dtpHoraInicio.Value.TimeOfDay);
        var fechaHoraFin    = _dtpFechaFin.Value.Date.Add(_dtpHoraFin.Value.TimeOfDay);

        if (fechaHoraInicio > fechaHoraFin)
        {
            MessageBox.Show(
                "La Fecha Inicial no puede ser mayor que la fecha final.",
                Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var parametros = new ReservaReporteParametros
            {
                FechaHoraInicio = fechaHoraInicio,
                FechaHoraFin    = fechaHoraFin,
                EstadoGenerado  = _chkGenerado.Checked,
                EstadoAtendido  = _chkAtendido.Checked,
                EstadoAnulado   = _chkAnulado.Checked,
                Orden           = (OrdenReserva)_cboOrden.SelectedIndex
            };

            var resultado = await _handler.HandleAsync(new ObtenerReporteReservasQuery(parametros));

            if (resultado.Filas.Count == 0)
            {
                _lblEstado.Text = "Sin resultados.";
                MessageBox.Show(
                    "No hay Datos para Mostrar.",
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var viewer = new ReporteViewer(
                Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);
            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("Reservas", resultado.Filas);
            viewer.Mostrar();

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} reservas";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte de reservas");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show(
                $"Error al generar el reporte: {ex.Message}",
                Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private void SetBotones(bool enabled)
    {
        foreach (var b in Controls.OfType<Button>())
            b.Enabled = enabled;
    }

    private void InitializeComponent()
    {
        Text = "Reservas";
        Size = new Size(520, 320);
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
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        // Date/time pickers
        _dtpFechaInicio = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _dtpHoraInicio  = new DateTimePicker { Format = DateTimePickerFormat.Time, ShowUpDown = true, Value = DateTime.Today };
        _dtpFechaFin    = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _dtpHoraFin     = new DateTimePicker { Format = DateTimePickerFormat.Time, ShowUpDown = true, Value = DateTime.Today.AddDays(1).AddSeconds(-1) };

        // State checkboxes — Generado checked by default (Legacy Value=1 for chkEstado(0))
        _chkGenerado = new CheckBox { Text = "Generado", Checked = true, AutoSize = true };
        _chkAtendido = new CheckBox { Text = "Atendido",  Checked = false, AutoSize = true };
        _chkAnulado  = new CheckBox { Text = "Anulado",   Checked = false, AutoSize = true };

        // Sort criteria combo
        _cboOrden = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        _cboOrden.Items.AddRange(["Reserva", "Nombres", "Teléfono", "Fecha", "Pax", "Estado"]);
        _cboOrden.SelectedIndex = 0;

        _lblEstado = new Label { Dock = DockStyle.Fill, ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft };

        // Action buttons
        var acciones = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft
        };
        var btnSalir  = new Button { Text = "Salir",  Width = 80 };
        var btnEmitir = new Button { Text = "Emitir", Width = 80 };
        btnSalir.Click  += (_, _) => Close();
        btnEmitir.Click += BtnEmitir_Click;
        acciones.Controls.AddRange([btnSalir, btnEmitir]);

        // Layout: row 0 — Fecha inicial + Hora inicial
        panel.Controls.Add(new Label { Text = "Fecha inicial:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        panel.Controls.Add(_dtpFechaInicio, 1, 0);
        panel.Controls.Add(new Label { Text = "Hora inicial:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
        panel.Controls.Add(_dtpHoraInicio, 3, 0);

        // Row 1 — Fecha final + Hora final
        panel.Controls.Add(new Label { Text = "Fecha final:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        panel.Controls.Add(_dtpFechaFin, 1, 1);
        panel.Controls.Add(new Label { Text = "Hora final:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 1);
        panel.Controls.Add(_dtpHoraFin, 3, 1);

        // Row 2 — State checkboxes
        panel.Controls.Add(new Label { Text = "Estado:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        var estadoPanel = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
        estadoPanel.Controls.AddRange([_chkGenerado, _chkAtendido, _chkAnulado]);
        panel.Controls.Add(estadoPanel, 1, 2);
        panel.SetColumnSpan(estadoPanel, 3);

        // Row 3 — Sort criteria
        panel.Controls.Add(new Label { Text = "Ordenado por:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
        panel.Controls.Add(_cboOrden, 1, 3);
        panel.SetColumnSpan(_cboOrden, 3);

        // Row 5 — status + buttons
        panel.Controls.Add(_lblEstado, 0, 5);
        panel.SetColumnSpan(_lblEstado, 3);
        panel.Controls.Add(acciones, 3, 5);

        Controls.Add(panel);
    }
}
