using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Formulario de filtro y emisión del Reporte de Comanda.
/// <para>
/// Legacy: <c>frmRepComanda.frm</c> → <c>spRep_Comanda</c>
/// Crystal: <c>dsrComandaD.Dsr</c> (detallado), <c>dsrComandaR.Dsr</c> (resumido)
/// FastReport: <c>RepComandaDetallado.frx</c> (detallado), <c>RepComanda.frx</c> (resumido)
/// Regla: BR-REP-001
/// </para>
/// <remarks>
/// Filtros heredados del Legacy:
/// cboMozo / chkTodoMozo → <see cref="_chkTodosMeseros"/> / <see cref="_txtMozo"/>
/// cboUsuario / chkTodoUsuario → <see cref="_chkTodosUsuarios"/> / <see cref="_txtUsuario"/>
/// dtpFecIni + dtpHoraIni → <see cref="_dtpInicio"/> (formato dd/MM/yyyy HH:mm)
/// dtpFecFin + dtpHoraFin → <see cref="_dtpFin"/> (formato dd/MM/yyyy HH:mm)
/// OptDetalle / OptResumen → <see cref="_rdoDetallado"/> / <see cref="_rdoResumido"/>
/// cboOrden → <see cref="_cmbOrden"/>
/// </remarks>
/// </summary>
public sealed class FrmComandaReporte : Form
{
    private readonly ObtenerReporteComandaHandler _handler;
    private readonly ILogger<FrmComandaReporte> _logger;
    private readonly string _rutaPlantillas;

    private DateTimePicker _dtpInicio = null!;
    private DateTimePicker _dtpFin = null!;
    private ComboBox _cmbOrden = null!;
    private RadioButton _rdoDetallado = null!;
    private RadioButton _rdoResumido = null!;
    private CheckBox _chkTodosMeseros = null!;
    private TextBox _txtMozo = null!;
    private CheckBox _chkTodosUsuarios = null!;
    private TextBox _txtUsuario = null!;
    private Label _lblEstado = null!;

    public FrmComandaReporte(
        ObtenerReporteComandaHandler handler,
        ILogger<FrmComandaReporte> logger,
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

        if (!_chkTodosMeseros.Checked && string.IsNullOrWhiteSpace(_txtMozo.Text))
        {
            MessageBox.Show("Debe ingresar un Mesero o marcar 'Todos los Meseros'.", Text,
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_chkTodosUsuarios.Checked && string.IsNullOrWhiteSpace(_txtUsuario.Text))
        {
            MessageBox.Show("Debe ingresar un Usuario o marcar 'Todos los Usuarios'.", Text,
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetBotonesHabilitados(false);
            _lblEstado.Text = "Generando reporte...";

            var criterio = ConstruirCriterio();
            var query = new ObtenerReporteComandaQuery(
                FlagTipo: _rdoDetallado.Checked,
                Orden: _cmbOrden.SelectedItem?.ToString() ?? "PEDIDO",
                FechaInicio: _dtpInicio.Value,
                FechaFin: _dtpFin.Value,
                Criterio: criterio);

            var resultado = await _handler.HandleAsync(query);

            using var viewer = new ReporteViewer(_logger as ILogger<ReporteViewer>
                ?? Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                _rutaPlantillas);

            viewer.CargarPlantilla(resultado.NombrePlantilla);
            viewer.AgregarDataSource("Comanda", resultado.Filas);

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
            _logger.LogError(ex, "Error al generar reporte Comanda");
            MessageBox.Show($"Error al generar reporte: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            _lblEstado.Text = "Error al generar reporte";
        }
        finally
        {
            SetBotonesHabilitados(true);
        }
    }

    /// <summary>
    /// Construye el criterio SQL para <c>@sCriterio</c> del SP.
    /// Legacy: sCriterio &amp;= " Mozo = '...'" / "usuario = '...'"
    /// </summary>
    private string ConstruirCriterio()
    {
        var partes = new List<string>();

        if (!_chkTodosMeseros.Checked && !string.IsNullOrWhiteSpace(_txtMozo.Text))
            partes.Add($"Mozo = '{_txtMozo.Text.Trim().Replace("'", "''")}'");

        if (!_chkTodosUsuarios.Checked && !string.IsNullOrWhiteSpace(_txtUsuario.Text))
            partes.Add($"usuario = '{_txtUsuario.Text.Trim().Replace("'", "''")}'");

        return string.Join(" and ", partes);
    }

    private void SetBotonesHabilitados(bool habilitado)
    {
        foreach (Control c in Controls)
            if (c is Button btn && btn.Name is "BtnEmitir" or "BtnImprimir")
                btn.Enabled = habilitado;
    }

    private void InitializeComponent()
    {
        Text = "Reporte de Comanda";
        Size = new Size(680, 400);
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
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        // ── Fecha Inicio ──
        _dtpInicio = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "dd/MM/yyyy HH:mm",
            Value = DateTime.Today
        };
        panel.Controls.Add(new Label { Text = "Fecha Inicio:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 0);
        panel.Controls.Add(_dtpInicio, 1, 0);

        // ── Fecha Fin ──
        _dtpFin = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "dd/MM/yyyy HH:mm",
            Value = DateTime.Today.AddDays(1).AddSeconds(-1)
        };
        panel.Controls.Add(new Label { Text = "Fecha Fin:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 1);
        panel.Controls.Add(_dtpFin, 1, 1);

        // ── Mozo ──
        _chkTodosMeseros = new CheckBox { Text = "Todos los Meseros", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left };
        _txtMozo = new TextBox { Enabled = false, Dock = DockStyle.Fill };
        _chkTodosMeseros.CheckedChanged += (_, _) => _txtMozo.Enabled = !_chkTodosMeseros.Checked;
        panel.Controls.Add(_chkTodosMeseros, 0, 2);
        panel.Controls.Add(_txtMozo, 1, 2);

        // ── Usuario ──
        _chkTodosUsuarios = new CheckBox { Text = "Todos los Usuarios", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left };
        _txtUsuario = new TextBox { Enabled = false, Dock = DockStyle.Fill };
        _chkTodosUsuarios.CheckedChanged += (_, _) => _txtUsuario.Enabled = !_chkTodosUsuarios.Checked;
        panel.Controls.Add(_chkTodosUsuarios, 0, 3);
        panel.Controls.Add(_txtUsuario, 1, 3);

        // ── Ordenar por ──
        _cmbOrden = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        _cmbOrden.Items.AddRange(["PEDIDO", "COMANDA"]);
        _cmbOrden.SelectedIndex = 0;
        panel.Controls.Add(new Label { Text = "Ordenar por:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 4);
        panel.Controls.Add(_cmbOrden, 1, 4);

        // ── Tipo de Reporte ──
        _rdoDetallado = new RadioButton { Text = "Detallado", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left };
        _rdoResumido = new RadioButton { Text = "Resumido", AutoSize = true, Anchor = AnchorStyles.Left };
        var pnlTipo = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
        pnlTipo.Controls.Add(_rdoDetallado);
        pnlTipo.Controls.Add(_rdoResumido);
        panel.Controls.Add(new Label { Text = "Tipo Reporte:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 5);
        panel.Controls.Add(pnlTipo, 1, 5);

        // ── Estado ──
        _lblEstado = new Label { Dock = DockStyle.Fill, ForeColor = System.Drawing.Color.DimGray, AutoSize = false };
        panel.SetColumnSpan(_lblEstado, 2);
        panel.Controls.Add(_lblEstado, 0, 6);

        // ── Botones ──
        var pnlBotones = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            Dock = DockStyle.Fill
        };
        var btnEmitir = new Button { Text = "Ver Reporte", Width = 120, Name = "BtnEmitir" };
        var btnImprimir = new Button { Text = "Imprimir", Width = 100, Name = "BtnImprimir" };
        var btnCerrar = new Button { Text = "Cerrar", Width = 100 };
        btnEmitir.Click += BtnEmitir_Click;
        btnImprimir.Click += BtnImprimir_Click;
        btnCerrar.Click += (_, _) => Close();
        pnlBotones.Controls.AddRange([btnEmitir, btnImprimir, btnCerrar]);
        panel.SetColumnSpan(pnlBotones, 2);
        panel.Controls.Add(pnlBotones, 0, 7);

        Controls.Add(panel);
    }
}
