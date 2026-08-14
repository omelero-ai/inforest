using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte de Cuenta Corriente Integrado.
/// <para>
/// Legacy: <c>frmRepCtaCteIntegrado.frm</c> → <c>spRep_CtaCteIntegrado</c>
/// Crystal: <c>dsrRepCtaCteIntegradoConsolidado.Dsr</c>, <c>dsrRepCtaCteIntegradoDetallado.Dsr</c>
/// FastReport: <c>RepCtaCteIntegradoConsolidado.frx</c>, <c>RepCtaCteIntegradoDetallado.frx</c>
/// Regla: BR-REP-004
/// </para>
/// </summary>
public sealed class FrmCtaCteIntegradoReporte : Form
{
    private readonly ObtenerReporteCtaCteIntegradoHandler _handler;
    private readonly ILogger<FrmCtaCteIntegradoReporte> _logger;
    private readonly string _rutaPlantillas;

    public FrmCtaCteIntegradoReporte(
        ObtenerReporteCtaCteIntegradoHandler handler,
        ILogger<FrmCtaCteIntegradoReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        InitializeComponent();
    }

    private async void BtnVer_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: true, imprimir: false);

    private async void BtnImprimir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: true);

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir)
    {
        if (DtpInicio.Value > DtpFin.Value)
        {
            MessageBox.Show(
                "La fecha inicial no puede ser mayor a la fecha final.",
                "Validación",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        try
        {
            BtnVer.Enabled = false;
            BtnImprimir.Enabled = false;
            lblEstado.Text = "Generando reporte...";

            var resultado = await _handler.HandleAsync(new ObtenerReporteCtaCteIntegradoQuery(
                ObtenerFlagTipoSeleccionado(),
                DtpInicio.Value.Date,
                DtpFin.Value.Date.AddDays(1).AddSeconds(-1)));

            using var viewer = new ReporteViewer(
                _logger as ILogger<ReporteViewer>
                ?? Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);

            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("CtaCteIntegrado", resultado.Filas);

            if (mostrar)
                viewer.Mostrar();
            else if (imprimir)
                viewer.Imprimir();

            lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} registros";
        }
        catch (FileNotFoundException ex)
        {
            MessageBox.Show(
                $"Plantilla no encontrada: {ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            lblEstado.Text = "Plantilla pendiente";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al generar reporte de cuenta corriente integrado");
            MessageBox.Show(
                $"Error al generar reporte: {ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            lblEstado.Text = "Error al generar reporte";
        }
        finally
        {
            BtnVer.Enabled = true;
            BtnImprimir.Enabled = true;
        }
    }

    private string ObtenerFlagTipoSeleccionado()
        => RbConsolidado.Checked ? "1" :
           RbDetallado.Checked ? "2" : "3";

    private void InitializeComponent()
    {
        Text = "Cuenta Corriente Integrado";
        Size = new Size(500, 280);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var lblInicio = new Label { Text = "Fecha Inicio:", Location = new Point(12, 20), AutoSize = true };
        DtpInicio = new DateTimePicker
        {
            Location = new Point(120, 16),
            Width = 140,
            Format = DateTimePickerFormat.Short,
            Value = DateTime.Today.AddDays(-3)
        };

        var lblFin = new Label { Text = "Fecha Fin:", Location = new Point(12, 56), AutoSize = true };
        DtpFin = new DateTimePicker
        {
            Location = new Point(120, 52),
            Width = 140,
            Format = DateTimePickerFormat.Short,
            Value = DateTime.Today
        };

        var grpTipo = new GroupBox
        {
            Text = "Tipo de reporte",
            Location = new Point(12, 92),
            Size = new Size(456, 64)
        };

        RbConsolidado = new RadioButton { Text = "Consolidado", Location = new Point(16, 28), AutoSize = true, Checked = true };
        RbDetallado = new RadioButton { Text = "Detallado", Location = new Point(156, 28), AutoSize = true };
        RbEstado = new RadioButton { Text = "Estado", Location = new Point(280, 28), AutoSize = true };
        grpTipo.Controls.AddRange([RbConsolidado, RbDetallado, RbEstado]);

        BtnVer = new Button { Text = "Ver Reporte", Location = new Point(120, 176), Width = 110 };
        BtnImprimir = new Button { Text = "Imprimir", Location = new Point(242, 176), Width = 90 };
        var btnCerrar = new Button { Text = "Cerrar", Location = new Point(344, 176), Width = 90 };
        btnCerrar.Click += (_, _) => Close();

        lblEstado = new Label
        {
            Location = new Point(12, 220),
            Width = 456,
            AutoSize = false,
            ForeColor = Color.Gray
        };

        BtnVer.Click += BtnVer_Click;
        BtnImprimir.Click += BtnImprimir_Click;

        Controls.AddRange([lblInicio, DtpInicio, lblFin, DtpFin, grpTipo, BtnVer, BtnImprimir, btnCerrar, lblEstado]);
    }

    private DateTimePicker DtpInicio = null!;
    private DateTimePicker DtpFin = null!;
    private RadioButton RbConsolidado = null!;
    private RadioButton RbDetallado = null!;
    private RadioButton RbEstado = null!;
    private Button BtnVer = null!;
    private Button BtnImprimir = null!;
    private Label lblEstado = null!;
}
