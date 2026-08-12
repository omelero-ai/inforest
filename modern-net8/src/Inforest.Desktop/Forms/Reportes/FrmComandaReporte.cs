using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte de Comanda.
/// <para>
/// Legacy: <c>frmRepComanda.frm</c> → <c>spRep_Comanda</c>
/// Crystal: <c>dsrComandaD.Dsr</c>, <c>dsrComandaR.Dsr</c>
/// FastReport: <c>RepComanda.frx</c>, <c>RepComandaDetallado.frx</c>
/// Regla: BR-REP-001
/// </para>
/// </summary>
public partial class FrmComandaReporte : Form
{
    private readonly ObtenerReporteComandaHandler _handler;
    private readonly ILogger<FrmComandaReporte> _logger;
    private readonly string _rutaPlantillas;

    public FrmComandaReporte(
        ObtenerReporteComandaHandler handler,
        ILogger<FrmComandaReporte> logger,
        string rutaPlantillas)
    {
        InitializeComponent();
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
    }

    private async void BtnVer_Click(object? sender, EventArgs e)
    {
        await EmitirReporteAsync(mostrar: true, imprimir: false);
    }

    private async void BtnImprimir_Click(object? sender, EventArgs e)
    {
        await EmitirReporteAsync(mostrar: false, imprimir: true);
    }

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir)
    {
        if (!ValidarFiltros()) return;

        try
        {
            BtnVer.Enabled = false;
            BtnImprimir.Enabled = false;
            lblEstado.Text = "Generando reporte...";

            var query = new ObtenerReporteComandaQuery(
                FlagTipo: ChkDetallado.Checked,
                Orden: CmbOrden.SelectedItem?.ToString() ?? string.Empty,
                FechaInicio: DtpInicio.Value.Date,
                FechaFin: DtpFin.Value.Date.AddDays(1).AddSeconds(-1),
                Criterio: TxtCriterio.Text.Trim());

            var resultado = await _handler.HandleAsync(query);

            using var viewer = new ReporteViewer(_logger as ILogger<ReporteViewer>
                ?? Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);

            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("Comanda", resultado.Filas);

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
            _logger.LogError(ex, "Error al generar reporte Comanda");
            MessageBox.Show($"Error al generar reporte: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            lblEstado.Text = "Error al generar reporte";
        }
        finally
        {
            BtnVer.Enabled = true;
            BtnImprimir.Enabled = true;
        }
    }

    private bool ValidarFiltros()
    {
        if (DtpInicio.Value > DtpFin.Value)
        {
            MessageBox.Show("La fecha inicio no puede ser mayor a la fecha fin.", "Validación",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
        return true;
    }

    private void InitializeComponent()
    {
        Text = "Reporte de Comanda";
        Size = new Size(520, 340);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        // ── Controles de filtro ──
        var lblInicio = new Label { Text = "Fecha Inicio:", Location = new Point(12, 20), AutoSize = true };
        DtpInicio = new DateTimePicker { Location = new Point(120, 16), Width = 160, Format = DateTimePickerFormat.Short, Value = DateTime.Today };

        var lblFin = new Label { Text = "Fecha Fin:", Location = new Point(12, 54), AutoSize = true };
        DtpFin = new DateTimePicker { Location = new Point(120, 50), Width = 160, Format = DateTimePickerFormat.Short, Value = DateTime.Today };

        var lblOrden = new Label { Text = "Ordenar por:", Location = new Point(12, 88), AutoSize = true };
        CmbOrden = new ComboBox { Location = new Point(120, 84), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
        CmbOrden.Items.AddRange(["Pedido", "Comanda", "Mozo", "Fecha"]);
        CmbOrden.SelectedIndex = 0;

        ChkDetallado = new CheckBox { Text = "Detallado (por ítem)", Location = new Point(120, 120), AutoSize = true };

        var lblCriterio = new Label { Text = "Filtro adicional:", Location = new Point(12, 156), AutoSize = true };
        TxtCriterio = new TextBox { Location = new Point(120, 152), Width = 350, PlaceholderText = "Condición SQL adicional (opcional)" };

        // ── Botones ──
        BtnVer = new Button { Text = "Ver Reporte", Location = new Point(120, 220), Width = 120 };
        BtnImprimir = new Button { Text = "Imprimir", Location = new Point(260, 220), Width = 100 };
        var BtnCerrar = new Button { Text = "Cerrar", Location = new Point(380, 220), Width = 100 };
        BtnCerrar.Click += (s, e) => Close();

        lblEstado = new Label { Location = new Point(12, 264), Width = 480, AutoSize = false, ForeColor = System.Drawing.Color.Gray };

        BtnVer.Click += BtnVer_Click;
        BtnImprimir.Click += BtnImprimir_Click;

        Controls.AddRange([lblInicio, DtpInicio, lblFin, DtpFin, lblOrden, CmbOrden,
            ChkDetallado, lblCriterio, TxtCriterio, BtnVer, BtnImprimir, BtnCerrar, lblEstado]);
    }

    // ── Campos de controles ──
    private DateTimePicker DtpInicio = null!;
    private DateTimePicker DtpFin = null!;
    private ComboBox CmbOrden = null!;
    private CheckBox ChkDetallado = null!;
    private TextBox TxtCriterio = null!;
    private Button BtnVer = null!;
    private Button BtnImprimir = null!;
    private Label lblEstado = null!;
}
