using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte de Propinas.
/// <para>
/// Legacy: <c>frmRepPropina.frm</c> → <c>spRep_Propina</c>
/// Crystal: <c>dsrPropinaD.Dsr</c>, <c>dsrPropinaR.Dsr</c>
/// FastReport: <c>RepPropina.frx</c>
/// Regla: BR-REP-002
/// </para>
/// </summary>
public partial class FrmPropinaReporte : Form
{
    private readonly ObtenerReportePropinaHandler _handler;
    private readonly ILogger<FrmPropinaReporte> _logger;
    private readonly string _rutaPlantillas;

    public FrmPropinaReporte(
        ObtenerReportePropinaHandler handler,
        ILogger<FrmPropinaReporte> logger,
        string rutaPlantillas)
    {
        InitializeComponent();
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
    }

    private async void BtnVer_Click(object? sender, EventArgs e) =>
        await EmitirReporteAsync(mostrar: true, imprimir: false);

    private async void BtnImprimir_Click(object? sender, EventArgs e) =>
        await EmitirReporteAsync(mostrar: false, imprimir: true);

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir)
    {
        if (DtpInicio.Value > DtpFin.Value)
        {
            MessageBox.Show("La fecha inicio no puede ser mayor a la fecha fin.", "Validación",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            BtnVer.Enabled = false;
            BtnImprimir.Enabled = false;
            lblEstado.Text = "Generando reporte...";

            var query = new ObtenerReportePropinaQuery(
                FechaInicio: DtpInicio.Value.Date,
                FechaFin: DtpFin.Value.Date.AddDays(1).AddSeconds(-1),
                Condicion: TxtCondicion.Text.Trim());

            var resultado = await _handler.HandleAsync(query);

            using var viewer = new ReporteViewer(
                Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);

            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("Propinas", resultado.Filas);

            if (mostrar) viewer.Mostrar();
            else if (imprimir) viewer.Imprimir();

            lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} registros";
        }
        catch (FileNotFoundException ex)
        {
            MessageBox.Show($"Plantilla no encontrada: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al generar reporte Propina");
            MessageBox.Show($"Error al generar reporte: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            BtnVer.Enabled = true;
            BtnImprimir.Enabled = true;
        }
    }

    private void InitializeComponent()
    {
        Text = "Reporte de Propinas";
        Size = new Size(480, 280);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var lblInicio = new Label { Text = "Fecha Inicio:", Location = new Point(12, 20), AutoSize = true };
        DtpInicio = new DateTimePicker { Location = new Point(120, 16), Width = 160, Format = DateTimePickerFormat.Short, Value = DateTime.Today };

        var lblFin = new Label { Text = "Fecha Fin:", Location = new Point(12, 54), AutoSize = true };
        DtpFin = new DateTimePicker { Location = new Point(120, 50), Width = 160, Format = DateTimePickerFormat.Short, Value = DateTime.Today };

        var lblCond = new Label { Text = "Condición:", Location = new Point(12, 88), AutoSize = true };
        TxtCondicion = new TextBox { Location = new Point(120, 84), Width = 320, PlaceholderText = "Filtro adicional (opcional)" };

        BtnVer = new Button { Text = "Ver Reporte", Location = new Point(120, 160), Width = 120 };
        BtnImprimir = new Button { Text = "Imprimir", Location = new Point(260, 160), Width = 100 };
        var BtnCerrar = new Button { Text = "Cerrar", Location = new Point(370, 160), Width = 80 };
        BtnCerrar.Click += (s, e) => Close();

        lblEstado = new Label { Location = new Point(12, 200), Width = 450, AutoSize = false, ForeColor = System.Drawing.Color.Gray };

        BtnVer.Click += BtnVer_Click;
        BtnImprimir.Click += BtnImprimir_Click;

        Controls.AddRange([lblInicio, DtpInicio, lblFin, DtpFin, lblCond, TxtCondicion,
            BtnVer, BtnImprimir, BtnCerrar, lblEstado]);
    }

    private DateTimePicker DtpInicio = null!;
    private DateTimePicker DtpFin = null!;
    private TextBox TxtCondicion = null!;
    private Button BtnVer = null!;
    private Button BtnImprimir = null!;
    private Label lblEstado = null!;
}
