using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte de Venta Mensual / Diaria Integrada.
/// <para>
/// Legacy: <c>frmRepVentaMensualIntegrado.frm</c>, <c>frmRepDiarioVentaIntegrado.frm</c>
///         → <c>spRep_VentaMensualIntegrado</c>
/// Crystal: <c>dsrVentaMensualIntegrada.Dsr</c>, <c>dsrVentaDiariaIntegrada.Dsr</c>
/// FastReport: <c>RepVentaMensualIntegrado.frx</c>, <c>RepVentaDiariaIntegrado.frx</c>
/// Regla: BR-REP-009
/// </para>
/// </summary>
public partial class FrmVentaMensualIntegradoReporte : Form
{
    private readonly ObtenerReporteVentaMensualIntegradoHandler _handler;
    private readonly ILogger<FrmVentaMensualIntegradoReporte> _logger;
    private readonly string _rutaPlantillas;

    public FrmVentaMensualIntegradoReporte(
        ObtenerReporteVentaMensualIntegradoHandler handler,
        ILogger<FrmVentaMensualIntegradoReporte> logger,
        string rutaPlantillas)
    {
        InitializeComponent();
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        CargarAniosMeses();
    }

    private void CargarAniosMeses()
    {
        for (int y = DateTime.Today.Year; y >= DateTime.Today.Year - 3; y--)
            CmbAnio.Items.Add(y.ToString());
        CmbAnio.SelectedIndex = 0;

        for (int m = 1; m <= 12; m++)
            CmbMes.Items.Add(m.ToString("00"));
        CmbMes.SelectedIndex = DateTime.Today.Month - 1;
    }

    private async void BtnVer_Click(object? sender, EventArgs e) =>
        await EmitirReporteAsync(mostrar: true, imprimir: false);

    private async void BtnImprimir_Click(object? sender, EventArgs e) =>
        await EmitirReporteAsync(mostrar: false, imprimir: true);

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir)
    {
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

            var parametros = new VentaMensualIntegradoParametros
            {
                TipoPrecio = tipoPrecio,
                Ano = CmbAnio.SelectedItem?.ToString() ?? DateTime.Today.Year.ToString(),
                Mes = CmbMes.SelectedItem?.ToString() ?? DateTime.Today.Month.ToString("00"),
                HoraCorte = (double)NudHoraCorte.Value,
                Tipo = RbDiario.Checked ? "D" : "M",
                MetaMensual = (double)NudMeta.Value,
                DiasDelMes = DateTime.DaysInMonth(
                    int.Parse(CmbAnio.SelectedItem?.ToString() ?? DateTime.Today.Year.ToString()),
                    int.Parse(CmbMes.SelectedItem?.ToString() ?? DateTime.Today.Month.ToString()))
            };

            var resultado = await _handler.HandleAsync(new ObtenerReporteVentaMensualIntegradoQuery(parametros));

            using var viewer = new ReporteViewer(
                Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);
            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("VentaMensual", resultado.Filas);
            if (mostrar) viewer.Mostrar(); else viewer.Imprimir();

            lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} registros";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al generar reporte VentaMensualIntegrado");
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { BtnVer.Enabled = true; }
    }

    private void InitializeComponent()
    {
        Text = "Venta Mensual / Diaria Integrada";
        Size = new Size(500, 360);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        RbDiario = new RadioButton { Text = "Diario", Location = new Point(12, 16), AutoSize = true, Checked = true };
        RbMensual = new RadioButton { Text = "Mensual", Location = new Point(100, 16), AutoSize = true };

        var lblAnio = new Label { Text = "Año:", Location = new Point(12, 52), AutoSize = true };
        CmbAnio = new ComboBox { Location = new Point(100, 48), Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
        var lblMes = new Label { Text = "Mes:", Location = new Point(200, 52), AutoSize = true };
        CmbMes = new ComboBox { Location = new Point(260, 48), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };

        var lblPrecio = new Label { Text = "Tipo Precio:", Location = new Point(12, 88), AutoSize = true };
        CmbTipoPrecio = new ComboBox { Location = new Point(100, 84), Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
        CmbTipoPrecio.Items.AddRange(["Venta", "Neto", "Costo"]);
        CmbTipoPrecio.SelectedIndex = 0;

        var lblHora = new Label { Text = "Hora de Corte:", Location = new Point(12, 124), AutoSize = true };
        NudHoraCorte = new NumericUpDown { Location = new Point(120, 120), Width = 60, Minimum = 0, Maximum = 23, Value = 6 };

        var lblMeta = new Label { Text = "Meta Mensual:", Location = new Point(12, 160), AutoSize = true };
        NudMeta = new NumericUpDown { Location = new Point(120, 156), Width = 120, Maximum = 9999999, DecimalPlaces = 2 };

        BtnVer = new Button { Text = "Ver Reporte", Location = new Point(100, 230), Width = 120 };
        BtnImprimir = new Button { Text = "Imprimir", Location = new Point(240, 230), Width = 100 };
        var BtnCerrar = new Button { Text = "Cerrar", Location = new Point(360, 230), Width = 80 };
        BtnCerrar.Click += (s, e) => Close();

        lblEstado = new Label { Location = new Point(12, 278), Width = 460, AutoSize = false, ForeColor = System.Drawing.Color.Gray };

        BtnVer.Click += BtnVer_Click;
        BtnImprimir.Click += BtnImprimir_Click;

        Controls.AddRange([RbDiario, RbMensual, lblAnio, CmbAnio, lblMes, CmbMes,
            lblPrecio, CmbTipoPrecio, lblHora, NudHoraCorte, lblMeta, NudMeta,
            BtnVer, BtnImprimir, BtnCerrar, lblEstado]);
    }

    private RadioButton RbDiario = null!, RbMensual = null!;
    private ComboBox CmbAnio = null!, CmbMes = null!, CmbTipoPrecio = null!;
    private NumericUpDown NudHoraCorte = null!, NudMeta = null!;
    private Button BtnVer = null!, BtnImprimir = null!;
    private Label lblEstado = null!;
}
