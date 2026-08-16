using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte Liquidación de Cajero — filtros por turno, rango de fechas, usuario y sector de venta.
/// Legacy: <c>frmRepLiquidacion.frm</c> (InfoRest.vbp) — "Liquidación de Cajero"
/// Regla: BR-REP-021
/// SQL: <c>spRep_LiquidacionOutPut</c>, <c>spRep_Liquidacion</c> (tipos 1-5), <c>spRep_LiquidacionSuma</c>
/// Tablas: MDOCUMENTO, DPAGODOCUMENTO, MPEDIDO, MTURNO, TTARJETACREDITO, vTipoCancelacion,
///         vSectorVenta, TUSUARIO
/// </summary>
public sealed class FrmRepLiquidacionReporte : Form
{
    private readonly ObtenerReporteLiquidacionHandler _handler;
    private readonly IReporteRepository _repo;
    private readonly ILogger<FrmRepLiquidacionReporte> _logger;
    private readonly string _rutaPlantillas;

    // ── Controles ──────────────────────────────────────────────────────────────
    private RadioButton _rdoPorTurno = null!;
    private RadioButton _rdoPorFecha = null!;
    private TextBox _txtTurno = null!;
    private Button _btnBuscarTurno = null!;
    private DateTimePicker _dtpFecIni = null!;
    private DateTimePicker _dtpHorIni = null!;
    private DateTimePicker _dtpFecFin = null!;
    private DateTimePicker _dtpHorFin = null!;
    private CheckBox _chkDiaContable = null!;
    private CheckBox _chkTodosUsuarios = null!;
    private ComboBox _cboUsuario = null!;
    private CheckBox _chkTodosSectores = null!;
    private ComboBox _cboSectorVenta = null!;
    private CheckBox _chkCortesia = null!;
    private Label _lblEstado = null!;

    public FrmRepLiquidacionReporte(
        ObtenerReporteLiquidacionHandler handler,
        IReporteRepository repo,
        ILogger<FrmRepLiquidacionReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _repo = repo;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        InitializeComponent();
    }

    // ── Events ─────────────────────────────────────────────────────────────────

    private async void FrmRepLiquidacionReporte_Load(object? sender, EventArgs e)
        => await CargarCombosAsync();

    private void RdoPorTurno_CheckedChanged(object? sender, EventArgs e)
        => ActualizarEstadoControles();

    private void ChkTodosUsuarios_CheckedChanged(object? sender, EventArgs e)
        => _cboUsuario.Enabled = !_chkTodosUsuarios.Checked;

    private void ChkTodosSectores_CheckedChanged(object? sender, EventArgs e)
        => _cboSectorVenta.Enabled = !_chkTodosSectores.Checked;

    private void ChkDiaContable_CheckedChanged(object? sender, EventArgs e)
        => ActualizarEstadoControles();

    private void BtnBuscarTurno_Click(object? sender, EventArgs e)
    {
        // Validación simple: el usuario puede ingresar el código de turno manualmente
        // Legacy: frmBusca.Show vbModal con query MTURNO
        _txtTurno.Focus();
    }

    private async void BtnEmitir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: true, imprimir: false, exportarCsv: false);

    private async void BtnImprimir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: true, exportarCsv: false);

    private async void BtnExportar_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: false, exportarCsv: true);

    private void BtnCerrar_Click(object? sender, EventArgs e) => Close();

    // ── Cargar catálogos ───────────────────────────────────────────────────────

    private async Task CargarCombosAsync()
    {
        try
        {
            var usuariosTask = _repo.ObtenerUsuariosActivosAsync();
            var sectoresTask = _repo.ObtenerSectoresVentaAsync();
            await Task.WhenAll(usuariosTask, sectoresTask);

            _cboUsuario.Items.Clear();
            foreach (var u in usuariosTask.Result)
                _cboUsuario.Items.Add(new FiltroItem(u.Codigo, u.Descripcion));

            _cboSectorVenta.Items.Clear();
            foreach (var s in sectoresTask.Result)
                _cboSectorVenta.Items.Add(new FiltroItem(s.Codigo, s.Descripcion));
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "No se pudieron cargar los combos");
        }
    }

    // ── Emitir reporte ─────────────────────────────────────────────────────────

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir, bool exportarCsv)
    {
        if (!ValidarParametros(out var parametros)) return;

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var resultado = await _handler.HandleAsync(new ObtenerReporteLiquidacionQuery(parametros!));

            if (resultado.Documentos.Count == 0 && resultado.SumasGrupo.Count == 0)
            {
                _lblEstado.Text = "Sin resultados.";
                MessageBox.Show("No hay Datos para Mostrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (exportarCsv)
            {
                ExportarCsv(resultado);
            }
            else
            {
                using var viewer = new ReporteViewer(
                    Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                    _rutaPlantillas);
                viewer.CargarPlantilla("RepLiquidacion.frx");
                viewer.AgregarDataSource("Documentos", resultado.Documentos);
                viewer.AgregarDataSource("SumasGrupo", resultado.SumasGrupo);
                viewer.AgregarDataSource("Tarjetas", resultado.Tarjetas);
                viewer.AgregarDataSource("TiposPedido", resultado.TiposPedido);
                viewer.AgregarDataSource("OtrosTipos", resultado.OtrosTipos);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado — {resultado.Documentos.Count} documentos";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte Liquidación de Cajero");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private bool ValidarParametros(out LiquidacionParametros? parametros)
    {
        parametros = null;

        if (_rdoPorTurno.Checked && string.IsNullOrWhiteSpace(_txtTurno.Text))
        {
            MessageBox.Show("No se ha asignado el Turno.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (_rdoPorFecha.Checked && !_chkDiaContable.Checked)
        {
            var inicio = _dtpFecIni.Value.Date + _dtpHorIni.Value.TimeOfDay;
            var fin = _dtpFecFin.Value.Date + _dtpHorFin.Value.TimeOfDay;
            if (inicio > fin)
            {
                MessageBox.Show("Error en Rango de Fechas.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        if (!_chkTodosUsuarios.Checked && _cboUsuario.SelectedItem == null)
        {
            MessageBox.Show("No se ha asignado el Usuario.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (!_chkTodosSectores.Checked && _cboSectorVenta.SelectedItem == null)
        {
            MessageBox.Show("No se ha asignado el Sector de Venta.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        var fechaInicio = _chkDiaContable.Checked
            ? _dtpFecIni.Value.Date
            : _dtpFecIni.Value.Date + _dtpHorIni.Value.TimeOfDay;
        var fechaFin = _chkDiaContable.Checked
            ? _dtpFecFin.Value.Date
            : _dtpFecFin.Value.Date + _dtpHorFin.Value.TimeOfDay;

        parametros = new LiquidacionParametros
        {
            ModoFiltro = _rdoPorTurno.Checked
                ? LiquidacionModoFiltro.PorTurno
                : LiquidacionModoFiltro.PorFecha,
            Turno = _txtTurno.Text.Trim(),
            FechaInicio = fechaInicio,
            FechaFin = fechaFin,
            PorDiaContable = _chkDiaContable.Checked,
            Usuario = _chkTodosUsuarios.Checked
                ? string.Empty
                : (_cboUsuario.SelectedItem as FiltroItem)?.Codigo ?? string.Empty,
            SectorVenta = _chkTodosSectores.Checked
                ? string.Empty
                : (_cboSectorVenta.SelectedItem as FiltroItem)?.Codigo ?? string.Empty,
            IncluirCortesia = _chkCortesia.Checked,
            MostrarTodos = true
        };
        return true;
    }

    private static void ExportarCsv(LiquidacionResultado resultado)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"Liquidacion_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dialog.ShowDialog() != DialogResult.OK) return;

        var lineas = new List<string> { "Grupo,SubGrupo,Documento,Usuario,Fecha,NVenta1,NVenta2" };
        lineas.AddRange(resultado.Documentos.Select(d =>
            $"{d.Grupo},{d.SubGrupo},{d.TDocumento},{d.TUsuario},{d.FFecha},{d.NVenta1:0.00},{d.NVenta2:0.00}"));

        File.WriteAllLines(dialog.FileName, lineas, System.Text.Encoding.UTF8);
        MessageBox.Show($"Exportado: {dialog.FileName}", "Exportar CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ActualizarEstadoControles()
    {
        bool porTurno = _rdoPorTurno.Checked;
        _txtTurno.Enabled = porTurno;
        _btnBuscarTurno.Enabled = porTurno;
        _dtpFecIni.Enabled = !porTurno;
        _dtpHorIni.Enabled = !porTurno && !_chkDiaContable.Checked;
        _dtpFecFin.Enabled = !porTurno;
        _dtpHorFin.Enabled = !porTurno && !_chkDiaContable.Checked;
        _chkDiaContable.Enabled = !porTurno;
        _chkTodosSectores.Enabled = !porTurno;
        _cboSectorVenta.Enabled = !porTurno && !_chkTodosSectores.Checked;
    }

    private void SetBotones(bool habilitados)
    {
        foreach (Button btn in Controls.OfType<Button>())
            btn.Enabled = habilitados;
        foreach (Button btn in Controls.OfType<Panel>().SelectMany(p => p.Controls.OfType<Button>()))
            btn.Enabled = habilitados;
    }

    // ── InitializeComponent ────────────────────────────────────────────────────

    private void InitializeComponent()
    {
        Text = "Liquidación de Cajero";
        Size = new Size(640, 560);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9f);

        // ── Grupo Filtros ─────────────────────────────────────────────────────
        var grpFiltros = new GroupBox
        {
            Text = "Filtros",
            Location = new Point(12, 8),
            Size = new Size(602, 460),
            BackColor = Color.White,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold)
        };

        // Modo: Turno / Fecha
        var lblModo = new Label { Text = "Filtrar por:", Location = new Point(12, 26), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _rdoPorTurno = new RadioButton
        {
            Text = "Turno específico",
            Location = new Point(140, 22),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };
        _rdoPorFecha = new RadioButton
        {
            Text = "Rango de fechas",
            Location = new Point(290, 22),
            AutoSize = true,
            Checked = true,
            Font = new Font("Segoe UI", 9f)
        };

        // Turno
        var lblTurno = new Label { Text = "Turno:", Location = new Point(12, 58), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _txtTurno = new TextBox { Location = new Point(140, 55), Width = 140, Enabled = false, Font = new Font("Segoe UI", 9f) };
        _btnBuscarTurno = new Button
        {
            Text = "...",
            Location = new Point(286, 53),
            Size = new Size(30, 24),
            Enabled = false,
            Font = new Font("Segoe UI", 9f)
        };

        // Fechas
        var lblFecIni = new Label { Text = "Fecha inicio:", Location = new Point(12, 92), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _dtpFecIni = new DateTimePicker
        {
            Location = new Point(140, 88),
            Width = 140,
            Format = DateTimePickerFormat.Short,
            Font = new Font("Segoe UI", 9f),
            Value = DateTime.Today
        };
        _dtpHorIni = new DateTimePicker
        {
            Location = new Point(286, 88),
            Width = 100,
            Format = DateTimePickerFormat.Time,
            ShowUpDown = true,
            Font = new Font("Segoe UI", 9f),
            Value = DateTime.Today
        };

        var lblFecFin = new Label { Text = "Fecha fin:", Location = new Point(12, 124), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _dtpFecFin = new DateTimePicker
        {
            Location = new Point(140, 120),
            Width = 140,
            Format = DateTimePickerFormat.Short,
            Font = new Font("Segoe UI", 9f),
            Value = DateTime.Today
        };
        _dtpHorFin = new DateTimePicker
        {
            Location = new Point(286, 120),
            Width = 100,
            Format = DateTimePickerFormat.Time,
            ShowUpDown = true,
            Font = new Font("Segoe UI", 9f),
            Value = DateTime.Now
        };

        // Día contable
        _chkDiaContable = new CheckBox
        {
            Text = "Por Día Contable (solo fecha, sin hora)",
            Location = new Point(140, 153),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };

        // Usuario
        var lblUsuario = new Label { Text = "Usuario:", Location = new Point(12, 185), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _chkTodosUsuarios = new CheckBox
        {
            Text = "Todos",
            Location = new Point(140, 182),
            AutoSize = true,
            Checked = true,
            Font = new Font("Segoe UI", 9f)
        };
        _cboUsuario = new ComboBox
        {
            Location = new Point(230, 181),
            Width = 200,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Enabled = false,
            Font = new Font("Segoe UI", 9f)
        };

        // Sector de venta
        var lblSector = new Label { Text = "Sector venta:", Location = new Point(12, 220), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _chkTodosSectores = new CheckBox
        {
            Text = "Todos",
            Location = new Point(140, 217),
            AutoSize = true,
            Checked = true,
            Font = new Font("Segoe UI", 9f)
        };
        _cboSectorVenta = new ComboBox
        {
            Location = new Point(230, 216),
            Width = 200,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Enabled = false,
            Font = new Font("Segoe UI", 9f)
        };

        // Cortesía
        _chkCortesia = new CheckBox
        {
            Text = "Incluir Cortesías en el cálculo",
            Location = new Point(12, 254),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };

        grpFiltros.Controls.AddRange(new Control[]
        {
            lblModo, _rdoPorTurno, _rdoPorFecha,
            lblTurno, _txtTurno, _btnBuscarTurno,
            lblFecIni, _dtpFecIni, _dtpHorIni,
            lblFecFin, _dtpFecFin, _dtpHorFin,
            _chkDiaContable,
            lblUsuario, _chkTodosUsuarios, _cboUsuario,
            lblSector, _chkTodosSectores, _cboSectorVenta,
            _chkCortesia
        });

        // ── Botones ───────────────────────────────────────────────────────────
        var btnEmitir = new Button
        {
            Text = "Vista Previa",
            Location = new Point(12, 476),
            Size = new Size(120, 36),
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnEmitir.FlatAppearance.BorderSize = 0;

        var btnImprimir = new Button
        {
            Text = "Imprimir",
            Location = new Point(144, 476),
            Size = new Size(110, 36),
            BackColor = Color.FromArgb(16, 124, 16),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnImprimir.FlatAppearance.BorderSize = 0;

        var btnExportar = new Button
        {
            Text = "Exportar CSV",
            Location = new Point(266, 476),
            Size = new Size(120, 36),
            BackColor = Color.FromArgb(136, 0, 21),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnExportar.FlatAppearance.BorderSize = 0;

        var btnCerrar = new Button
        {
            Text = "Cerrar",
            Location = new Point(494, 476),
            Size = new Size(120, 36),
            BackColor = Color.FromArgb(96, 96, 96),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnCerrar.FlatAppearance.BorderSize = 0;

        // ── Estado ────────────────────────────────────────────────────────────
        _lblEstado = new Label
        {
            Text = string.Empty,
            Location = new Point(12, 524),
            Size = new Size(602, 20),
            ForeColor = Color.FromArgb(0, 120, 212),
            Font = new Font("Segoe UI", 8.5f)
        };

        // ── Wire events ───────────────────────────────────────────────────────
        _rdoPorTurno.CheckedChanged += RdoPorTurno_CheckedChanged;
        _rdoPorFecha.CheckedChanged += RdoPorTurno_CheckedChanged;
        _chkTodosUsuarios.CheckedChanged += ChkTodosUsuarios_CheckedChanged;
        _chkTodosSectores.CheckedChanged += ChkTodosSectores_CheckedChanged;
        _chkDiaContable.CheckedChanged += ChkDiaContable_CheckedChanged;
        _btnBuscarTurno.Click += BtnBuscarTurno_Click;
        btnEmitir.Click += BtnEmitir_Click;
        btnImprimir.Click += BtnImprimir_Click;
        btnExportar.Click += BtnExportar_Click;
        btnCerrar.Click += BtnCerrar_Click;
        Load += FrmRepLiquidacionReporte_Load;

        Controls.AddRange(new Control[]
        {
            grpFiltros,
            btnEmitir, btnImprimir, btnExportar, btnCerrar,
            _lblEstado
        });
    }

    // ── Helper ─────────────────────────────────────────────────────────────────

    private sealed class FiltroItem
    {
        public string Codigo { get; }
        private readonly string _descripcion;
        public FiltroItem(string codigo, string descripcion) { Codigo = codigo; _descripcion = descripcion; }
        public override string ToString() => _descripcion;
    }
}
