using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte de Propinas.
/// <para>
/// Legacy: <c>frmRepPropina.frm</c> → <c>spRep_Propina</c><br/>
/// Crystal: <c>dsrPropinaD.Dsr</c> (Detallado), <c>dsrPropinaR.Dsr</c> (Resumido)<br/>
/// FastReport: <c>RepPropina.frx</c> (Detallado), <c>RepPropinaResumido.frx</c> (Resumido)<br/>
/// Regla: BR-REP-002
/// </para>
/// <remarks>
/// Filtros equivalentes al Legacy:<br/>
/// <c>dtpFecIni + dtpHorIni</c> → <see cref="_dtpInicio"/> (formato dd/MM/yyyy HH:mm)<br/>
/// <c>dtpFecFin + dtpHorFin</c> → <see cref="_dtpFin"/> (formato dd/MM/yyyy HH:mm)<br/>
/// <c>cboMozo + chkMozo</c>             → <see cref="_chkTodosMozos"/> / <see cref="_txtMozo"/><br/>
/// <c>cboMotorizado + chkMotorizado</c>  → <see cref="_chkTodosMotorizados"/> / <see cref="_txtMotorizado"/><br/>
/// <c>cboTpedido + chkTpedido</c>       → <see cref="_chkTodosTipoPedido"/> / <see cref="_txtTipoPedido"/><br/>
/// <c>optOpcion(0)=Detallado / optOpcion(1)=Resumido</c> → <see cref="_rdoDetallado"/> / <see cref="_rdoResumido"/>
/// </remarks>
/// </summary>
public sealed class FrmPropinaReporte : Form
{
    private readonly ObtenerReportePropinaHandler _handler;
    private readonly ILogger<FrmPropinaReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private CheckBox _chkTodosMozos = null!;
    private TextBox _txtMozo = null!;
    private CheckBox _chkTodosMotorizados = null!;
    private TextBox _txtMotorizado = null!;
    private CheckBox _chkTodosTipoPedido = null!;
    private TextBox _txtTipoPedido = null!;
    private RadioButton _rdoDetallado = null!;
    private RadioButton _rdoResumido = null!;
    private Label _lblEstado = null!;

    public FrmPropinaReporte(
        ObtenerReportePropinaHandler handler,
        ILogger<FrmPropinaReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        InitializeComponent();
    }

    private async void BtnEmitir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: true, imprimir: false);

    private async void BtnImprimir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: true);

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir)
    {
        if (_dtpInicio.Value > _dtpFin.Value)
        {
            MessageBox.Show("Error en Rango de Fechas.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosMozos.Checked && string.IsNullOrWhiteSpace(_txtMozo.Text))
        {
            MessageBox.Show("Debe escoger un Mesero.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosMotorizados.Checked && string.IsNullOrWhiteSpace(_txtMotorizado.Text))
        {
            MessageBox.Show("Debe escoger un Motorizado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosTipoPedido.Checked && string.IsNullOrWhiteSpace(_txtTipoPedido.Text))
        {
            MessageBox.Show("Debe escoger un Tipo de Pedido.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var condicion = ConstruirCondicion();
            var query = new ObtenerReportePropinaQuery(
                FechaInicio: _dtpInicio.Value,
                FechaFin: _dtpFin.Value,
                Condicion: condicion,
                EsDetallado: _rdoDetallado.Checked);

            var resultado = await _handler.HandleAsync(query);

            if (resultado.Filas.Count == 0)
            {
                _lblEstado.Text = "Sin resultados.";
                MessageBox.Show("No hay Datos para Mostrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var viewer = new ReporteViewer(
                Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);

            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("Propinas", resultado.Filas);

            if (mostrar) viewer.Mostrar();
            else if (imprimir) viewer.Imprimir();

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} registros";
        }
        catch (FileNotFoundException ex)
        {
            MessageBox.Show($"Plantilla no encontrada: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al generar reporte Propina");
            MessageBox.Show($"Error al generar reporte: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            _lblEstado.Text = "Error al generar reporte";
        }
        finally
        {
            SetBotones(true);
        }
    }

    /// <summary>
    /// Construye la condición SQL pasada a <c>@sCondicion</c> del SP.
    /// Equivalente al Sub <c>Genera</c> de <c>frmRepPropina.frm</c>:
    /// <c>tmozo = '...'</c>, <c>mpedido.tmotorizado = '...'</c>, <c>mpedido.ttipopedido = '...'</c>
    /// </summary>
    private string ConstruirCondicion()
    {
        var partes = new List<string>();

        if (!_chkTodosMozos.Checked && !string.IsNullOrWhiteSpace(_txtMozo.Text))
            partes.Add($"tmozo = '{_txtMozo.Text.Trim().Replace("'", "''")}'");

        if (!_chkTodosMotorizados.Checked && !string.IsNullOrWhiteSpace(_txtMotorizado.Text))
            partes.Add($"mpedido.tmotorizado = '{_txtMotorizado.Text.Trim().Replace("'", "''")}'");

        if (!_chkTodosTipoPedido.Checked && !string.IsNullOrWhiteSpace(_txtTipoPedido.Text))
            partes.Add($"mpedido.ttipopedido = '{_txtTipoPedido.Text.Trim().Replace("'", "''")}'");

        return string.Join(" and ", partes);
    }

    private void SetBotones(bool enabled)
    {
        foreach (var btn in Controls.OfType<Panel>()
            .SelectMany(p => p.Controls.OfType<Button>())
            .Concat(Controls.OfType<Button>()))
            btn.Enabled = enabled;
    }

    private void InitializeComponent()
    {
        Text = "Reporte de Propinas";
        Size = new Size(640, 420);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 9,
            Padding = new Padding(12)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 210));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (var i = 0; i < 9; i++)
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // ── Fecha Inicio ──────────────────────────────────────────────────────
        _dtpInicio = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "dd/MM/yyyy HH:mm",
            Value = DateTime.Today
        };
        panel.Controls.Add(new Label { Text = "Fecha Inicio:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 0);
        panel.Controls.Add(_dtpInicio, 1, 0);

        // ── Fecha Fin ─────────────────────────────────────────────────────────
        _dtpFin = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "dd/MM/yyyy HH:mm",
            Value = DateTime.Today.AddDays(1).AddSeconds(-1)
        };
        panel.Controls.Add(new Label { Text = "Fecha Fin:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 1);
        panel.Controls.Add(_dtpFin, 1, 1);

        // ── Mozo (cboMozo + chkMozo) ─────────────────────────────────────────
        _chkTodosMozos = new CheckBox { Text = "Todos los Meseros", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left };
        _txtMozo = new TextBox { Enabled = false, Dock = DockStyle.Fill, PlaceholderText = "Código mesero" };
        _chkTodosMozos.CheckedChanged += (_, _) =>
        {
            _txtMozo.Enabled = !_chkTodosMozos.Checked;
            if (_chkTodosMozos.Checked) _txtMozo.Text = string.Empty;
        };
        panel.Controls.Add(_chkTodosMozos, 0, 2);
        panel.Controls.Add(_txtMozo, 1, 2);

        // ── Motorizado (cboMotorizado + chkMotorizado) ────────────────────────
        _chkTodosMotorizados = new CheckBox { Text = "Todos los Motorizados", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left };
        _txtMotorizado = new TextBox { Enabled = false, Dock = DockStyle.Fill, PlaceholderText = "Código motorizado" };
        _chkTodosMotorizados.CheckedChanged += (_, _) =>
        {
            _txtMotorizado.Enabled = !_chkTodosMotorizados.Checked;
            if (_chkTodosMotorizados.Checked) _txtMotorizado.Text = string.Empty;
        };
        panel.Controls.Add(_chkTodosMotorizados, 0, 3);
        panel.Controls.Add(_txtMotorizado, 1, 3);

        // ── Tipo Pedido (cboTpedido + chkTpedido) ────────────────────────────
        _chkTodosTipoPedido = new CheckBox { Text = "Todos los Tipos de Pedido", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left };
        _txtTipoPedido = new TextBox { Enabled = false, Dock = DockStyle.Fill, PlaceholderText = "Código tipo pedido" };
        _chkTodosTipoPedido.CheckedChanged += (_, _) =>
        {
            _txtTipoPedido.Enabled = !_chkTodosTipoPedido.Checked;
            if (_chkTodosTipoPedido.Checked) _txtTipoPedido.Text = string.Empty;
        };
        panel.Controls.Add(_chkTodosTipoPedido, 0, 4);
        panel.Controls.Add(_txtTipoPedido, 1, 4);

        // ── Tipo de Reporte (optOpcion 0=Detallado, 1=Resumido) ───────────────
        _rdoDetallado = new RadioButton { Text = "Detallado", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left };
        _rdoResumido = new RadioButton { Text = "Resumido", AutoSize = true, Anchor = AnchorStyles.Left };
        var pnlTipo = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
        pnlTipo.Controls.Add(_rdoDetallado);
        pnlTipo.Controls.Add(_rdoResumido);
        panel.Controls.Add(new Label { Text = "Tipo Reporte:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 5);
        panel.Controls.Add(pnlTipo, 1, 5);

        // ── Estado ────────────────────────────────────────────────────────────
        _lblEstado = new Label { Dock = DockStyle.Fill, ForeColor = Color.DimGray, AutoSize = false };
        panel.SetColumnSpan(_lblEstado, 2);
        panel.Controls.Add(_lblEstado, 0, 6);

        // ── Botones ───────────────────────────────────────────────────────────
        var pnlBotones = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            Dock = DockStyle.Fill,
            AutoSize = true
        };
        var btnCerrar = new Button { Text = "Cerrar", Width = 90 };
        var btnImprimir = new Button { Text = "Imprimir", Width = 90 };
        var btnEmitir = new Button { Text = "Ver Reporte", Width = 110 };
        btnCerrar.Click += (_, _) => Close();
        btnImprimir.Click += BtnImprimir_Click;
        btnEmitir.Click += BtnEmitir_Click;
        pnlBotones.Controls.AddRange([btnCerrar, btnImprimir, btnEmitir]);
        panel.SetColumnSpan(pnlBotones, 2);
        panel.Controls.Add(pnlBotones, 0, 7);

        Controls.Add(panel);
    }
}
