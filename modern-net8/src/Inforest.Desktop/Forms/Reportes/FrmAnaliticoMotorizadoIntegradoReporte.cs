using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte Analítico de Motorizado Integrado.
/// <para>
/// Legacy: <c>frmRepAnaliticoMotorizadoIntegrado.frm</c> → <c>spRep_AnaliticoMotorizadoIntegrado</c>
/// Crystal: <c>dsrAnaliticoMotorizadoIntegradoConsolidado.Dsr</c>
/// FastReport: <c>RepAnaliticoMotorizadoIntegrado.frx</c>
/// Regla: BR-REP-010
/// </para>
/// </summary>
public partial class FrmAnaliticoMotorizadoIntegradoReporte : Form
{
    private readonly ObtenerReporteAnaliticoMotorizadoIntegradoHandler _handler;
    private readonly ILogger<FrmAnaliticoMotorizadoIntegradoReporte> _logger;
    private readonly string _rutaPlantillas;

    public FrmAnaliticoMotorizadoIntegradoReporte(
        ObtenerReporteAnaliticoMotorizadoIntegradoHandler handler,
        ILogger<FrmAnaliticoMotorizadoIntegradoReporte> logger,
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
            lblEstado.Text = "Generando reporte...";

            var tipoPrecio = CmbTipoPrecio.SelectedIndex switch
            {
                1 => ExpresionPrecio.Neto,
                2 => ExpresionPrecio.Costo,
                _ => ExpresionPrecio.Venta
            };

            var parametros = new AnaliticoMotorizadoIntegradoParametros
            {
                FlagProduccion = ChkProduccion.Checked,
                FlagVenta = ChkVenta.Checked,
                FlagCortesia = ChkCortesia.Checked,
                FlagCuentaCte = ChkCtaCte.Checked,
                FlagPedidosFacturados = ChkFacturados.Checked,
                FlagCombinacion = ChkCombinacion.Checked,
                FlagCargo = ChkCargo.Checked,
                TipoPrecio = tipoPrecio,
                FechaInicio = DtpInicio.Value.Date,
                FechaFin = DtpFin.Value.Date.AddDays(1).AddSeconds(-1)
            };

            var resultado = await _handler.HandleAsync(
                new ObtenerReporteAnaliticoMotorizadoIntegradoQuery(parametros));

            using var viewer = new ReporteViewer(
                Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);
            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("AnaliticoMotorizado", resultado.Filas);
            if (mostrar) viewer.Mostrar(); else viewer.Imprimir();

            lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} registros";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al generar reporte AnaliticoMotorizadoIntegrado");
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { BtnVer.Enabled = true; }
    }

    private void InitializeComponent()
    {
        Text = "Analítico Motorizado Integrado";
        Size = new Size(520, 400);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var lblInicio = new Label { Text = "Fecha Inicio:", Location = new Point(12, 16), AutoSize = true };
        DtpInicio = new DateTimePicker { Location = new Point(120, 12), Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        var lblFin = new Label { Text = "Fecha Fin:", Location = new Point(290, 16), AutoSize = true };
        DtpFin = new DateTimePicker { Location = new Point(370, 12), Width = 130, Format = DateTimePickerFormat.Short, Value = DateTime.Today };

        var grp = new GroupBox { Text = "Incluir en el reporte", Location = new Point(12, 50), Size = new Size(490, 120) };
        ChkProduccion = new CheckBox { Text = "Producción", Location = new Point(12, 24), AutoSize = true, Checked = true };
        ChkVenta = new CheckBox { Text = "Venta", Location = new Point(120, 24), AutoSize = true, Checked = true };
        ChkCortesia = new CheckBox { Text = "Cortesía", Location = new Point(200, 24), AutoSize = true };
        ChkCtaCte = new CheckBox { Text = "Cta.Cte.", Location = new Point(290, 24), AutoSize = true };
        ChkFacturados = new CheckBox { Text = "Solo Facturados", Location = new Point(12, 56), AutoSize = true };
        ChkCombinacion = new CheckBox { Text = "Combinación", Location = new Point(160, 56), AutoSize = true };
        ChkCargo = new CheckBox { Text = "Con Cargo", Location = new Point(290, 56), AutoSize = true };
        grp.Controls.AddRange([ChkProduccion, ChkVenta, ChkCortesia, ChkCtaCte, ChkFacturados, ChkCombinacion, ChkCargo]);

        var lblPrecio = new Label { Text = "Tipo Precio:", Location = new Point(12, 184), AutoSize = true };
        CmbTipoPrecio = new ComboBox { Location = new Point(120, 180), Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
        CmbTipoPrecio.Items.AddRange(["Venta", "Neto", "Costo"]);
        CmbTipoPrecio.SelectedIndex = 0;

        BtnVer = new Button { Text = "Ver Reporte", Location = new Point(120, 280), Width = 120 };
        BtnImprimir = new Button { Text = "Imprimir", Location = new Point(260, 280), Width = 100 };
        var BtnCerrar = new Button { Text = "Cerrar", Location = new Point(400, 280), Width = 80 };
        BtnCerrar.Click += (s, e) => Close();

        lblEstado = new Label { Location = new Point(12, 330), Width = 490, AutoSize = false, ForeColor = System.Drawing.Color.Gray };

        BtnVer.Click += BtnVer_Click;
        BtnImprimir.Click += BtnImprimir_Click;

        Controls.AddRange([lblInicio, DtpInicio, lblFin, DtpFin, grp,
            lblPrecio, CmbTipoPrecio, BtnVer, BtnImprimir, BtnCerrar, lblEstado]);
    }

    private DateTimePicker DtpInicio = null!, DtpFin = null!;
    private CheckBox ChkProduccion = null!, ChkVenta = null!, ChkCortesia = null!,
        ChkCtaCte = null!, ChkFacturados = null!, ChkCombinacion = null!, ChkCargo = null!;
    private ComboBox CmbTipoPrecio = null!;
    private Button BtnVer = null!, BtnImprimir = null!;
    private Label lblEstado = null!;
}
