using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte de Tiempo KDS.
/// <para>
/// Legacy: KDS monitor → <c>USP_KDS_ResporteTiempoPedido</c>, <c>USP_KDS_ResporteTiempoProducto</c>
/// Crystal: <c>dsrTiempoKdsP.Dsr</c>, <c>dsrTiempoKdsR.Dsr</c>
/// FastReport: <c>RepTiempoKdsPedido.frx</c>, <c>RepTiempoKdsProducto.frx</c>
/// Reglas: BR-REP-011, BR-REP-012
/// </para>
/// </summary>
public partial class FrmTiempoKdsReporte : Form
{
    private readonly ObtenerReporteTiempoKdsPedidoHandler _handlerPedido;
    private readonly ObtenerReporteTiempoKdsProductoHandler _handlerProducto;
    private readonly ILogger<FrmTiempoKdsReporte> _logger;
    private readonly string _rutaPlantillas;

    public FrmTiempoKdsReporte(
        ObtenerReporteTiempoKdsPedidoHandler handlerPedido,
        ObtenerReporteTiempoKdsProductoHandler handlerProducto,
        ILogger<FrmTiempoKdsReporte> logger,
        string rutaPlantillas)
    {
        InitializeComponent();
        _handlerPedido = handlerPedido;
        _handlerProducto = handlerProducto;
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
            lblEstado.Text = "Generando reporte KDS...";

            using var viewer = new ReporteViewer(
                Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);

            if (RbPorPedido.Checked)
            {
                var query = new ObtenerReporteTiempoKdsPedidoQuery(DtpInicio.Value.Date, DtpFin.Value.Date.AddDays(1).AddSeconds(-1));
                var resultado = await _handlerPedido.HandleAsync(query);
                viewer.CargarPlantilla(resultado.NombrePlantilla);
                viewer.AgregarDataSource("TiempoKdsPedido", resultado.Filas);
                lblEstado.Text = $"Reporte KDS por Pedido: {resultado.Filas.Count} registros";
            }
            else
            {
                var query = new ObtenerReporteTiempoKdsProductoQuery(
                    DtpInicio.Value.Date,
                    DtpFin.Value.Date.AddDays(1).AddSeconds(-1),
                    TxtGrupo.Text.Trim(),
                    TxtSubGrupo.Text.Trim(),
                    TxtProducto.Text.Trim());
                var resultado = await _handlerProducto.HandleAsync(query);
                viewer.CargarPlantilla(resultado.NombrePlantilla);
                viewer.AgregarDataSource("TiempoKdsProducto", resultado.Filas);
                lblEstado.Text = $"Reporte KDS por Producto: {resultado.Filas.Count} registros";
            }

            if (mostrar) viewer.Mostrar(); else viewer.Imprimir();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al generar reporte TiempoKDS");
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { BtnVer.Enabled = true; }
    }

    private void RbPorPedido_CheckedChanged(object? sender, EventArgs e)
    {
        TxtGrupo.Enabled = RbPorProducto.Checked;
        TxtSubGrupo.Enabled = RbPorProducto.Checked;
        TxtProducto.Enabled = RbPorProducto.Checked;
    }

    private void InitializeComponent()
    {
        Text = "Tiempo KDS";
        Size = new Size(480, 360);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        RbPorPedido = new RadioButton { Text = "Por Pedido", Location = new Point(12, 16), AutoSize = true, Checked = true };
        RbPorProducto = new RadioButton { Text = "Por Producto", Location = new Point(120, 16), AutoSize = true };
        RbPorPedido.CheckedChanged += RbPorPedido_CheckedChanged;

        var lblInicio = new Label { Text = "Fecha Inicio:", Location = new Point(12, 52), AutoSize = true };
        DtpInicio = new DateTimePicker { Location = new Point(120, 48), Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        var lblFin = new Label { Text = "Fecha Fin:", Location = new Point(12, 86), AutoSize = true };
        DtpFin = new DateTimePicker { Location = new Point(120, 82), Width = 150, Format = DateTimePickerFormat.Short, Value = DateTime.Today };

        var lblGrupo = new Label { Text = "Grupo:", Location = new Point(12, 122), AutoSize = true };
        TxtGrupo = new TextBox { Location = new Point(120, 118), Width = 140, Enabled = false, PlaceholderText = "Todos" };
        var lblSubGrupo = new Label { Text = "Sub-Grupo:", Location = new Point(12, 156), AutoSize = true };
        TxtSubGrupo = new TextBox { Location = new Point(120, 152), Width = 140, Enabled = false, PlaceholderText = "Todos" };
        var lblProducto = new Label { Text = "Producto:", Location = new Point(12, 190), AutoSize = true };
        TxtProducto = new TextBox { Location = new Point(120, 186), Width = 200, Enabled = false, PlaceholderText = "Todos" };

        BtnVer = new Button { Text = "Ver Reporte", Location = new Point(100, 260), Width = 120 };
        BtnImprimir = new Button { Text = "Imprimir", Location = new Point(240, 260), Width = 100 };
        var BtnCerrar = new Button { Text = "Cerrar", Location = new Point(360, 260), Width = 80 };
        BtnCerrar.Click += (s, e) => Close();

        lblEstado = new Label { Location = new Point(12, 306), Width = 450, AutoSize = false, ForeColor = System.Drawing.Color.Gray };

        BtnVer.Click += BtnVer_Click;
        BtnImprimir.Click += BtnImprimir_Click;

        Controls.AddRange([RbPorPedido, RbPorProducto, lblInicio, DtpInicio, lblFin, DtpFin,
            lblGrupo, TxtGrupo, lblSubGrupo, TxtSubGrupo, lblProducto, TxtProducto,
            BtnVer, BtnImprimir, BtnCerrar, lblEstado]);
    }

    private RadioButton RbPorPedido = null!, RbPorProducto = null!;
    private DateTimePicker DtpInicio = null!, DtpFin = null!;
    private TextBox TxtGrupo = null!, TxtSubGrupo = null!, TxtProducto = null!;
    private Button BtnVer = null!, BtnImprimir = null!;
    private Label lblEstado = null!;
}
