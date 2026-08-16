using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte comparativo de venta mensual por día (Salón / Delivery / Llevar / Canal4 / Canal5).
/// Legacy: <c>frmRepVentaFecha.frm</c> (InfoRest.vbp) — "Venta Mensual por Fechas"
/// Regla: BR-REP-020
/// SQL: <c>spRep_VentaFecha</c>
/// Tablas: MPEDIDO, DPEDIDO, MDOCUMENTO, vSubGrupo
/// </summary>
public sealed class FrmRepVentaFechaReporte : Form
{
    private readonly ObtenerReporteVentaFechaHandler _handler;
    private readonly IReporteRepository _repo;
    private readonly ILogger<FrmRepVentaFechaReporte> _logger;
    private readonly string _rutaPlantillas;

    // ── Controles ──────────────────────────────────────────────────────────────
    private NumericUpDown _nudAno = null!;
    private ComboBox _cboMes = null!;
    private NumericUpDown _nudHoraCierre = null!;
    private RadioButton _rdoPrecioVenta = null!;
    private RadioButton _rdoPrecioNeto = null!;
    private CheckBox _chkValorarPreventa = null!;
    private CheckBox _chkEvaluarPorDocumentos = null!;
    private RadioButton _rdoTodos = null!;
    private RadioButton _rdoSeleccion = null!;
    private CheckedListBox _clbSubGrupos = null!;
    private Label _lblEstado = null!;

    private static readonly string[] NombresMeses =
        ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"];

    public FrmRepVentaFechaReporte(
        ObtenerReporteVentaFechaHandler handler,
        IReporteRepository repo,
        ILogger<FrmRepVentaFechaReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _repo = repo;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        InitializeComponent();
    }

    // ── Events ─────────────────────────────────────────────────────────────────

    private async void FrmRepVentaFechaReporte_Load(object? sender, EventArgs e)
        => await CargarSubGruposAsync();

    private async void BtnEmitir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: true, imprimir: false, exportarCsv: false);

    private async void BtnImprimir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: true, exportarCsv: false);

    private async void BtnExportar_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: false, exportarCsv: true);

    private void BtnCerrar_Click(object? sender, EventArgs e) => Close();

    private void RdoTodos_CheckedChanged(object? sender, EventArgs e)
        => _clbSubGrupos.Enabled = !_rdoTodos.Checked;

    // ── Cargar catálogo de sub-grupos ──────────────────────────────────────────

    private async Task CargarSubGruposAsync()
    {
        try
        {
            var items = await _repo.ObtenerSubGruposAsync();
            _clbSubGrupos.Items.Clear();
            foreach (var sg in items)
                _clbSubGrupos.Items.Add(new SubGrupoListItem(sg.Codigo, sg.Descripcion), isChecked: true);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "No se pudieron cargar los sub-grupos");
        }
    }

    // ── Emitir reporte ─────────────────────────────────────────────────────────

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir, bool exportarCsv)
    {
        // Validar selección de sub-grupos cuando se usa filtro personalizado
        if (_rdoSeleccion.Checked)
        {
            var seleccionados = _clbSubGrupos.CheckedItems.Cast<SubGrupoListItem>().ToList();
            if (seleccionados.Count == 0)
            {
                MessageBox.Show("Debe escoger mínimo un subgrupo.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var subGruposFiltro = _rdoSeleccion.Checked
                ? _clbSubGrupos.CheckedItems.Cast<SubGrupoListItem>().Select(i => i.Codigo).ToArray()
                : Array.Empty<string>();

            var parametros = new VentaFechaParametros
            {
                Ano = (int)_nudAno.Value,
                Mes = _cboMes.SelectedIndex + 1,
                HoraCierre = (int)_nudHoraCierre.Value,
                TipoPrecio = _rdoPrecioNeto.Checked
                    ? TipoPrecioVentaFecha.Neto
                    : TipoPrecioVentaFecha.Venta,
                ValorarPreventaEnCero = _chkValorarPreventa.Checked,
                EvaluarPorDocumentos = _chkEvaluarPorDocumentos.Checked,
                SubGruposFiltro = subGruposFiltro
            };

            var resultado = await _handler.HandleAsync(new ObtenerReporteVentaFechaQuery(parametros));

            if (resultado.Filas.Count == 0)
            {
                _lblEstado.Text = "Sin resultados.";
                MessageBox.Show("No hay Datos para Mostrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (exportarCsv)
            {
                ExportarCsv(resultado.Filas, resultado.TituloReporte);
            }
            else
            {
                using var viewer = new ReporteViewer(
                    Microsoft.Extensions.Logging.Abstractions.NullLogger<ReporteViewer>.Instance,
                    _rutaPlantillas);
                viewer.CargarPlantilla(resultado.NombrePlantilla);
                viewer.AgregarDataSource("VentaFecha", resultado.Filas);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado: {resultado.Filas.Count} filas";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando reporte Venta Mensual por Fechas");
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private static void ExportarCsv(IReadOnlyList<VentaFechaRow> filas, string titulo)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"VentaMensual_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dialog.ShowDialog() != DialogResult.OK) return;

        var lineas = new List<string>
        {
            "Fecha,Dia,Salon,Delivery,Llevar,Canal4,Canal5,Venta,Cantidad,Pax"
        };
        lineas.AddRange(filas.Select(f => string.Join(",",
            f.Fecha.ToString("dd/MM/yyyy"),
            f.Dia,
            f.Salon.ToString("0.00"),
            f.Delivery.ToString("0.00"),
            f.Llevar.ToString("0.00"),
            f.Canal4.ToString("0.00"),
            f.Canal5.ToString("0.00"),
            f.Venta.ToString("0.00"),
            f.Cantidad,
            f.Pax)));

        File.WriteAllLines(dialog.FileName, lineas, System.Text.Encoding.UTF8);
        MessageBox.Show($"Exportado: {dialog.FileName}", "Exportar CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void SetBotones(bool habilitados)
    {
        foreach (Button btn in Controls.OfType<Button>())
            btn.Enabled = habilitados;
    }

    // ── InitializeComponent ────────────────────────────────────────────────────

    private void InitializeComponent()
    {
        Text = "Venta Mensual por Fechas";
        Size = new Size(620, 580);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9f);

        // ── Grupo Opciones ────────────────────────────────────────────────────
        var grpOpciones = new GroupBox
        {
            Text = "Opciones",
            Location = new Point(12, 8),
            Size = new Size(576, 460),
            BackColor = Color.White,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold)
        };

        // Año
        var lblAno = new Label { Text = "Año:", Location = new Point(12, 30), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _nudAno = new NumericUpDown
        {
            Location = new Point(160, 27),
            Width = 80,
            Minimum = 2000,
            Maximum = 2099,
            Value = DateTime.Today.Year,
            Font = new Font("Segoe UI", 9f)
        };

        // Mes
        var lblMes = new Label { Text = "Mes:", Location = new Point(12, 62), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _cboMes = new ComboBox
        {
            Location = new Point(160, 58),
            Width = 140,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 9f)
        };
        _cboMes.Items.AddRange(NombresMeses);
        _cboMes.SelectedIndex = DateTime.Today.Month - 1;

        // Hora de cierre
        var lblHora = new Label { Text = "Hora corte (HH):", Location = new Point(12, 94), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _nudHoraCierre = new NumericUpDown
        {
            Location = new Point(160, 90),
            Width = 60,
            Minimum = 0,
            Maximum = 23,
            Value = 0,
            Font = new Font("Segoe UI", 9f)
        };

        // Precio de venta/neto
        var lblPrecio = new Label { Text = "Tipo de precio:", Location = new Point(12, 126), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _rdoPrecioVenta = new RadioButton
        {
            Text = "Precio venta",
            Location = new Point(160, 122),
            AutoSize = true,
            Checked = true,
            Font = new Font("Segoe UI", 9f)
        };
        _rdoPrecioNeto = new RadioButton
        {
            Text = "Precio neto",
            Location = new Point(270, 122),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };

        // Checkboxes
        _chkValorarPreventa = new CheckBox
        {
            Text = "Valorizar con cero la pre venta facturada",
            Location = new Point(12, 156),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };
        _chkEvaluarPorDocumentos = new CheckBox
        {
            Text = "Evaluar la venta por documentos emitidos",
            Location = new Point(12, 183),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };

        // SubGrupos filter
        var lblSubGrupo = new Label { Text = "Sub-Grupos:", Location = new Point(12, 218), AutoSize = true, Font = new Font("Segoe UI", 9f) };
        _rdoTodos = new RadioButton
        {
            Text = "Todos",
            Location = new Point(160, 214),
            AutoSize = true,
            Checked = true,
            Font = new Font("Segoe UI", 9f)
        };
        _rdoSeleccion = new RadioButton
        {
            Text = "Selección",
            Location = new Point(240, 214),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };
        _clbSubGrupos = new CheckedListBox
        {
            Location = new Point(12, 240),
            Size = new Size(544, 200),
            Enabled = false,
            Font = new Font("Segoe UI", 9f),
            CheckOnClick = true
        };

        grpOpciones.Controls.AddRange(new Control[]
        {
            lblAno, _nudAno,
            lblMes, _cboMes,
            lblHora, _nudHoraCierre,
            lblPrecio, _rdoPrecioVenta, _rdoPrecioNeto,
            _chkValorarPreventa, _chkEvaluarPorDocumentos,
            lblSubGrupo, _rdoTodos, _rdoSeleccion, _clbSubGrupos
        });

        // ── Botones ───────────────────────────────────────────────────────────
        var btnEmitir = new Button
        {
            Text = "Vista Previa",
            Location = new Point(12, 476),
            Size = new Size(110, 36),
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
            Location = new Point(134, 476),
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
            Location = new Point(256, 476),
            Size = new Size(110, 36),
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
            Location = new Point(478, 476),
            Size = new Size(110, 36),
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
            Size = new Size(576, 20),
            ForeColor = Color.FromArgb(0, 120, 212),
            Font = new Font("Segoe UI", 8.5f)
        };

        // ── Wire events ───────────────────────────────────────────────────────
        btnEmitir.Click  += BtnEmitir_Click;
        btnImprimir.Click += BtnImprimir_Click;
        btnExportar.Click += BtnExportar_Click;
        btnCerrar.Click  += BtnCerrar_Click;
        _rdoTodos.CheckedChanged += RdoTodos_CheckedChanged;
        Load += FrmRepVentaFechaReporte_Load;

        Controls.AddRange(new Control[]
        {
            grpOpciones,
            btnEmitir, btnImprimir, btnExportar, btnCerrar,
            _lblEstado
        });
    }

    // ── Helper ─────────────────────────────────────────────────────────────────

    /// <summary>Encapsula un item de sub-grupo para CheckedListBox.</summary>
    private sealed class SubGrupoListItem
    {
        public string Codigo { get; }
        private readonly string _descripcion;
        public SubGrupoListItem(string codigo, string descripcion)
        {
            Codigo = codigo;
            _descripcion = descripcion;
        }
        public override string ToString() => _descripcion;
    }
}
