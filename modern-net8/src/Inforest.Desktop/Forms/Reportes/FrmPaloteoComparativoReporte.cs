using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte de Paloteo Comparativo.
/// <para>
/// Legacy: <c>frmRepPaloteoComparativo.frm</c> → <c>spRep_PaloteoComparativo</c>
/// Crystal: <c>dsrPaloteoComparativo.Dsr</c>
/// FastReport: <c>RepPaloteoComparativo.frx</c>
/// Regla: BR-REP-005
/// </para>
/// </summary>
public partial class FrmPaloteoComparativoReporte : Form
{
    private readonly ObtenerReportePaloteoComparativoHandler _handler;
    private readonly ILogger<FrmPaloteoComparativoReporte> _logger;
    private readonly string _rutaPlantillas;

    public FrmPaloteoComparativoReporte(
        ObtenerReportePaloteoComparativoHandler handler,
        ILogger<FrmPaloteoComparativoReporte> logger,
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

            var query = new ObtenerReportePaloteoComparativoQuery(
                PorTurno: RbTurno.Checked,
                PorValor: ChkPorValor.Checked,
                TipoProducto: TxtTipoProducto.Text.Trim(),
                AreaProduccion: TxtArea.Text.Trim(),
                Grupo: TxtGrupo.Text.Trim(),
                SubGrupo: TxtSubGrupo.Text.Trim(),
                CodigoProducto: TxtProducto.Text.Trim(),
                Turno: TxtTurno.Text.Trim(),
                FechaInicio: DtpInicio.Value.Date,
                FechaFin: DtpFin.Value.Date.AddDays(1).AddSeconds(-1),
                SoloFacturados: ChkSoloFacturados.Checked);

            var resultado = await _handler.HandleAsync(query);

            using var viewer = new ReporteViewer(
                Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);
            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("PaloteoComparativo", resultado.Filas);
            if (mostrar) viewer.Mostrar(); else viewer.Imprimir();

            lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} registros";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al generar reporte PaloteoComparativo");
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { BtnVer.Enabled = true; }
    }

    private void RbFecha_CheckedChanged(object? sender, EventArgs e)
    {
        DtpInicio.Enabled = RbFecha.Checked;
        DtpFin.Enabled = RbFecha.Checked;
        TxtTurno.Enabled = RbTurno.Checked;
    }

    private void InitializeComponent()
    {
        Text = "Paloteo Comparativo";
        Size = new Size(560, 420);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        RbFecha = new RadioButton { Text = "Por Fecha", Location = new Point(12, 16), AutoSize = true, Checked = true };
        RbTurno = new RadioButton { Text = "Por Turno", Location = new Point(120, 16), AutoSize = true };
        RbFecha.CheckedChanged += (s, e) => RbFecha_CheckedChanged(s, e);

        var lblInicio = new Label { Text = "Fecha Inicio:", Location = new Point(12, 48), AutoSize = true };
        DtpInicio = new DateTimePicker { Location = new Point(120, 44), Width = 140, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        var lblFin = new Label { Text = "Fecha Fin:", Location = new Point(280, 48), AutoSize = true };
        DtpFin = new DateTimePicker { Location = new Point(360, 44), Width = 140, Format = DateTimePickerFormat.Short, Value = DateTime.Today };

        var lblTurno = new Label { Text = "Turno:", Location = new Point(12, 82), AutoSize = true };
        TxtTurno = new TextBox { Location = new Point(120, 78), Width = 140, Enabled = false };

        var lblGrupo = new Label { Text = "Grupo:", Location = new Point(12, 116), AutoSize = true };
        TxtGrupo = new TextBox { Location = new Point(120, 112), Width = 140, PlaceholderText = "Todos" };
        var lblSubGrupo = new Label { Text = "Sub-Grupo:", Location = new Point(280, 116), AutoSize = true };
        TxtSubGrupo = new TextBox { Location = new Point(360, 112), Width = 140, PlaceholderText = "Todos" };

        var lblTipoProd = new Label { Text = "Tipo Prod.:", Location = new Point(12, 150), AutoSize = true };
        TxtTipoProducto = new TextBox { Location = new Point(120, 146), Width = 140, PlaceholderText = "Todos" };
        var lblArea = new Label { Text = "Área:", Location = new Point(280, 150), AutoSize = true };
        TxtArea = new TextBox { Location = new Point(360, 146), Width = 140, PlaceholderText = "Todos" };

        var lblProducto = new Label { Text = "Producto:", Location = new Point(12, 184), AutoSize = true };
        TxtProducto = new TextBox { Location = new Point(120, 180), Width = 200, PlaceholderText = "Todos" };

        ChkPorValor = new CheckBox { Text = "Mostrar por Valor", Location = new Point(12, 218), AutoSize = true };
        ChkSoloFacturados = new CheckBox { Text = "Solo Facturados", Location = new Point(200, 218), AutoSize = true };

        BtnVer = new Button { Text = "Ver Reporte", Location = new Point(120, 280), Width = 120 };
        BtnImprimir = new Button { Text = "Imprimir", Location = new Point(260, 280), Width = 100 };
        var BtnCerrar = new Button { Text = "Cerrar", Location = new Point(420, 280), Width = 80 };
        BtnCerrar.Click += (s, e) => Close();

        lblEstado = new Label { Location = new Point(12, 330), Width = 520, AutoSize = false, ForeColor = System.Drawing.Color.Gray };

        BtnVer.Click += BtnVer_Click;
        BtnImprimir.Click += BtnImprimir_Click;

        Controls.AddRange([RbFecha, RbTurno, lblInicio, DtpInicio, lblFin, DtpFin,
            lblTurno, TxtTurno, lblGrupo, TxtGrupo, lblSubGrupo, TxtSubGrupo,
            lblTipoProd, TxtTipoProducto, lblArea, TxtArea, lblProducto, TxtProducto,
            ChkPorValor, ChkSoloFacturados, BtnVer, BtnImprimir, BtnCerrar, lblEstado]);
    }

    private RadioButton RbFecha = null!, RbTurno = null!;
    private DateTimePicker DtpInicio = null!, DtpFin = null!;
    private TextBox TxtTurno = null!, TxtGrupo = null!, TxtSubGrupo = null!,
        TxtTipoProducto = null!, TxtArea = null!, TxtProducto = null!;
    private CheckBox ChkPorValor = null!, ChkSoloFacturados = null!;
    private Button BtnVer = null!, BtnImprimir = null!;
    private Label lblEstado = null!;
}
