using Inforest.Application.Reportes;
using Inforest.Desktop.Reportes;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Forms.Reportes;

/// <summary>
/// Reporte de Registro de Ventas — equivalente a <c>frmRepRegistroVenta.frm</c> de InfoRest.vbp.
/// Permite generar 8 tipos de reporte distintos sobre los documentos emitidos:
/// SUNAT, Estado de Documentos, Agrupado por Fechas, Agrupado por Tipo Doc,
/// Correlativo, Detallado por Comprobante, Correlativo Detallado y Con Forma de Pago.
/// Legacy: <c>frmRepRegistroVenta.frm</c> (InfoRest.vbp)
/// SQL: <c>spRep_RegVenta</c>, <c>spRep_RegVentaSunat</c>, <c>spRep_RegVentaSunatAD</c>, <c>spRep_ComprobanteDetallado</c>
/// Regla: BR-REP-022
/// </summary>
public sealed class FrmRepRegistroVentaReporte : Form
{
    private readonly ObtenerReporteRegistroVentaHandler _handler;
    private readonly IReporteRepository _repo;
    private readonly ILogger<FrmRepRegistroVentaReporte> _logger;
    private readonly string _rutaPlantillas;

    // ── Tipo de reporte ────────────────────────────────────────────────────────
    private RadioButton _rdoSunat = null!;
    private RadioButton _rdoEstado = null!;
    private RadioButton _rdoAgrupadoFechas = null!;
    private RadioButton _rdoAgrupadoTipoDoc = null!;
    private RadioButton _rdoCorrelativo = null!;
    private RadioButton _rdoDetalladoComprobante = null!;
    private RadioButton _rdoCorrelativoDetallado = null!;
    private RadioButton _rdoConFormaPago = null!;

    // ── Filtros ────────────────────────────────────────────────────────────────
    private ComboBox _cboTipoDocumento = null!;
    private ComboBox _cboEstado = null!;
    private ComboBox _cboCaja = null!;
    private ComboBox _cboTipoPago = null!;
    private ComboBox _cboOrden = null!;
    private ComboBox _cboRedondeo = null!;

    private CheckBox _chkTodosDoc = null!;
    private CheckBox _chkTodosEstado = null!;
    private CheckBox _chkTodasCajas = null!;
    private CheckBox _chkTodosPagos = null!;

    private TextBox _txtCliente = null!;
    private Button _btnBuscarCliente = null!;
    private string _codigoCliente = string.Empty;

    // ── Fechas ─────────────────────────────────────────────────────────────────
    private DateTimePicker _dtpFecIni = null!;
    private DateTimePicker _dtpHorIni = null!;
    private DateTimePicker _dtpFecFin = null!;
    private DateTimePicker _dtpHorFin = null!;

    // ── Modo Año/Mes (solo para AgrupadoPorFechas) ─────────────────────────────
    private Panel _pnlAnioMes = null!;
    private NumericUpDown _nudAno = null!;
    private ComboBox _cboMes = null!;
    private NumericUpDown _nudHoraCorte = null!;

    // ── Opciones extra ─────────────────────────────────────────────────────────
    private CheckBox _chkDiaContable = null!;
    private CheckBox _chkRegistroVenta = null!;
    private CheckBox _chkTransferenciaGratuita = null!;

    private Label _lblEstado = null!;

    private static readonly string[] NombresMeses =
        ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"];

    public FrmRepRegistroVentaReporte(
        ObtenerReporteRegistroVentaHandler handler,
        IReporteRepository repo,
        ILogger<FrmRepRegistroVentaReporte> logger,
        string rutaPlantillas = "Reports/Templates")
    {
        _handler = handler;
        _repo = repo;
        _logger = logger;
        _rutaPlantillas = rutaPlantillas;
        InitializeComponent();
    }

    // ── Eventos ────────────────────────────────────────────────────────────────

    private async void FrmRepRegistroVentaReporte_Load(object? sender, EventArgs e)
        => await CargarCombosAsync();

    private void RdoTipo_CheckedChanged(object? sender, EventArgs e)
    {
        if (sender is RadioButton rdo && rdo.Checked)
            ActualizarControlesSegunTipo();
    }

    private void ChkTodos_CheckedChanged(object? sender, EventArgs e)
    {
        if (sender == _chkTodosDoc)
        {
            _cboTipoDocumento.Enabled = !_chkTodosDoc.Checked;
            if (_chkTodosDoc.Checked) _cboTipoDocumento.SelectedIndex = -1;
        }
        else if (sender == _chkTodosEstado)
        {
            _cboEstado.Enabled = !_chkTodosEstado.Checked;
            if (_chkTodosEstado.Checked) _cboEstado.SelectedIndex = -1;
        }
        else if (sender == _chkTodasCajas)
        {
            _cboCaja.Enabled = !_chkTodasCajas.Checked;
            if (_chkTodasCajas.Checked) _cboCaja.SelectedIndex = -1;
        }
        else if (sender == _chkTodosPagos)
        {
            _cboTipoPago.Enabled = !_chkTodosPagos.Checked;
            if (_chkTodosPagos.Checked) _cboTipoPago.SelectedIndex = -1;
        }
    }

    private void ChkDiaContable_CheckedChanged(object? sender, EventArgs e)
    {
        bool sinHora = _chkDiaContable.Checked;
        _dtpHorIni.Visible = !sinHora;
        _dtpHorFin.Visible = !sinHora;
    }

    private async void BtnEmitir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: true, imprimir: false, exportarCsv: false);

    private async void BtnImprimir_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: true, exportarCsv: false);

    private async void BtnExportar_Click(object? sender, EventArgs e)
        => await EmitirReporteAsync(mostrar: false, imprimir: false, exportarCsv: true);

    private void BtnCerrar_Click(object? sender, EventArgs e) => Close();

    private void BtnBuscarCliente_Click(object? sender, EventArgs e)
    {
        // Sin frmBusca disponible en .NET 8: el usuario puede ingresar el código directamente
        _codigoCliente = _txtCliente.Text.Trim();
    }

    // ── Carga catálogos ────────────────────────────────────────────────────────

    private async Task CargarCombosAsync()
    {
        try
        {
            var tiposDocTask = _repo.ObtenerTiposDocumentoAsync();
            var estadosTask = _repo.ObtenerEstadosDocumentoAsync();
            var cajasTask = _repo.ObtenerCajasActivasAsync();
            var pagosTask = _repo.ObtenerTiposPagoAsync();

            await Task.WhenAll(tiposDocTask, estadosTask, cajasTask, pagosTask);

            CargarCombo(_cboTipoDocumento, tiposDocTask.Result);
            CargarCombo(_cboEstado, estadosTask.Result);
            CargarCombo(_cboCaja, cajasTask.Result);
            CargarCombo(_cboTipoPago, pagosTask.Result);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error cargando combos de filtro");
        }
    }

    private static void CargarCombo(ComboBox combo, IReadOnlyList<ReporteFiltroOpcion> items)
    {
        combo.Items.Clear();
        foreach (var item in items)
            combo.Items.Add(new FiltroItem(item.Codigo, item.Descripcion));
    }

    // ── Control de estado de controles según tipo ──────────────────────────────

    private void ActualizarControlesSegunTipo()
    {
        var tipo = TipoSeleccionado;

        // Modo Año/Mes solo para AgrupadoPorFechas
        _pnlAnioMes.Visible = tipo == TipoReporteRegistroVenta.AgrupadoPorFechas;

        // TipoPago: deshabilitado para SUNAT, AgrupadoTipoDoc, ConFormaPago
        bool tipoPagoDisponible = tipo != TipoReporteRegistroVenta.CorrelativoSunat
                                && tipo != TipoReporteRegistroVenta.AgrupadoPorTipoDocumento
                                && tipo != TipoReporteRegistroVenta.CorrelativoConFormaPago;
        _chkTodosPagos.Enabled = tipoPagoDisponible;
        _cboTipoPago.Enabled = tipoPagoDisponible && !_chkTodosPagos.Checked;
        if (!tipoPagoDisponible)
        {
            _chkTodosPagos.Checked = true;
            _cboTipoPago.SelectedIndex = -1;
        }

        // TransferenciaGratuita: solo para CorrelativoSunat
        _chkTransferenciaGratuita.Visible = tipo == TipoReporteRegistroVenta.CorrelativoSunat;

        // Orden y Redondeo: deshabilitados para DetalladoPorComprobante
        bool ordenDisponible = tipo != TipoReporteRegistroVenta.DetalladoPorComprobante;
        _cboOrden.Enabled = ordenDisponible;
        _cboRedondeo.Enabled = ordenDisponible;

        // ConFormaPago: GAP avisado
        if (tipo == TipoReporteRegistroVenta.CorrelativoConFormaPago)
            _lblEstado.Text = "GAP: spRep_RegVentaSunat_formaPago no encontrado en SQL Legacy.";
        else
            _lblEstado.Text = string.Empty;
    }

    private TipoReporteRegistroVenta TipoSeleccionado
    {
        get
        {
            if (_rdoSunat.Checked) return TipoReporteRegistroVenta.CorrelativoSunat;
            if (_rdoEstado.Checked) return TipoReporteRegistroVenta.EstadoDocumentos;
            if (_rdoAgrupadoFechas.Checked) return TipoReporteRegistroVenta.AgrupadoPorFechas;
            if (_rdoAgrupadoTipoDoc.Checked) return TipoReporteRegistroVenta.AgrupadoPorTipoDocumento;
            if (_rdoCorrelativo.Checked) return TipoReporteRegistroVenta.CorrelativoDocumento;
            if (_rdoDetalladoComprobante.Checked) return TipoReporteRegistroVenta.DetalladoPorComprobante;
            if (_rdoCorrelativoDetallado.Checked) return TipoReporteRegistroVenta.CorrelativoDetallado;
            return TipoReporteRegistroVenta.CorrelativoConFormaPago;
        }
    }

    // ── Emitir reporte ─────────────────────────────────────────────────────────

    private async Task EmitirReporteAsync(bool mostrar, bool imprimir, bool exportarCsv)
    {
        var tipo = TipoSeleccionado;

        if (tipo == TipoReporteRegistroVenta.CorrelativoConFormaPago)
        {
            MessageBox.Show(
                "El reporte 'Correlativo con Forma de Pago' no está disponible.\nGAP: spRep_RegVentaSunat_formaPago no encontrado en SQL Legacy.",
                Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Validar fechas
        if (_chkDiaContable.Checked)
        {
            if (_dtpFecIni.Value.Date > _dtpFecFin.Value.Date)
            {
                MessageBox.Show("Error en Rango de Fechas.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        else
        {
            var inicio = _dtpFecIni.Value.Date + _dtpHorIni.Value.TimeOfDay;
            var fin = _dtpFecFin.Value.Date + _dtpHorFin.Value.TimeOfDay;
            if (inicio > fin)
            {
                MessageBox.Show("Error en Rango de Fechas.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        try
        {
            SetBotones(false);
            _lblEstado.Text = "Generando reporte...";

            var parametros = ConstruirParametros(tipo);
            var resultado = await _handler.HandleAsync(new ObtenerReporteRegistroVentaQuery(parametros));

            if (resultado.TotalFilas == 0)
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
                viewer.CargarPlantilla(resultado.NombrePlantilla);
                AgregarDataSource(viewer, resultado);
                if (mostrar) viewer.Mostrar();
                else if (imprimir) viewer.Imprimir();
            }

            _lblEstado.Text = $"Reporte generado: {resultado.TotalFilas} filas";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generando Reporte Registro de Ventas tipo={Tipo}", tipo);
            _lblEstado.Text = "Error al generar el reporte.";
            MessageBox.Show($"Error al generar el reporte: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBotones(true);
        }
    }

    private RegistroVentaParametros ConstruirParametros(TipoReporteRegistroVenta tipo)
    {
        DateTime fIni, fFin;
        if (_chkDiaContable.Checked)
        {
            fIni = _dtpFecIni.Value.Date;
            fFin = _dtpFecFin.Value.Date.AddDays(1).AddSeconds(-1);
        }
        else
        {
            fIni = _dtpFecIni.Value.Date + _dtpHorIni.Value.TimeOfDay;
            fFin = _dtpFecFin.Value.Date + _dtpHorFin.Value.TimeOfDay;
        }

        string tipoDoc = _chkTodosDoc.Checked ? string.Empty
            : (_cboTipoDocumento.SelectedItem as FiltroItem)?.Codigo ?? string.Empty;
        string estadoDoc = _chkTodosEstado.Checked ? string.Empty
            : (_cboEstado.SelectedItem as FiltroItem)?.Codigo ?? string.Empty;
        string caja = _chkTodasCajas.Checked ? string.Empty
            : (_cboCaja.SelectedItem as FiltroItem)?.Codigo ?? string.Empty;
        string tipoPago = _chkTodosPagos.Checked ? string.Empty
            : (_cboTipoPago.SelectedItem as FiltroItem)?.Codigo ?? string.Empty;

        return new RegistroVentaParametros
        {
            TipoReporte = tipo,
            CodigoCliente = _codigoCliente,
            TipoDocumento = tipoDoc,
            EstadoDocumento = estadoDoc,
            Caja = caja,
            TipoPago = tipoPago,
            Orden = _cboOrden.Text,
            FechaInicio = fIni,
            FechaFin = fFin,
            DiaContable = _chkDiaContable.Checked,
            SoloRegistroVenta = _chkRegistroVenta.Checked,
            Redondeo = _cboRedondeo.Text.ToUpperInvariant().Contains("DOS"),
            TransferenciaGratuita = _chkTransferenciaGratuita.Checked,
            Ano = (int)_nudAno.Value,
            Mes = _cboMes.SelectedIndex + 1,
            HoraCorte = (int)_nudHoraCorte.Value
        };
    }

    private static void AgregarDataSource(ReporteViewer viewer, RegistroVentaResultado resultado)
    {
        switch (resultado.TipoReporte)
        {
            case TipoReporteRegistroVenta.CorrelativoSunat:
                viewer.AgregarDataSource("RegistroVentaSunat", resultado.FilasSunat);
                break;
            case TipoReporteRegistroVenta.AgrupadoPorTipoDocumento:
                viewer.AgregarDataSource("RegistroVentaSunatAd", resultado.FilasSunatAd);
                break;
            case TipoReporteRegistroVenta.DetalladoPorComprobante:
                viewer.AgregarDataSource("RegistroVentaDetallado", resultado.FilasDetallado);
                break;
            default:
                viewer.AgregarDataSource("RegistroVenta", resultado.Filas);
                break;
        }
    }

    private static void ExportarCsv(RegistroVentaResultado resultado)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"RegistroVentas_{resultado.TipoReporte}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dialog.ShowDialog() != DialogResult.OK) return;

        List<string> lineas;
        switch (resultado.TipoReporte)
        {
            case TipoReporteRegistroVenta.CorrelativoSunat:
                lineas = ["Voucher,FEmision,TDoc,Serie,Numero,RazonSocial,BaseImOpGra,Exonerada,Inafecta,IGV,OtrosTrib,ImporteT,TipoCambio,Descuento,Recargo"];
                lineas.AddRange(resultado.FilasSunat.Select(f => string.Join(",",
                    Csv(f.Voucher), f.FEmision.ToString("dd/MM/yyyy"), Csv(f.TDoc),
                    Csv(f.Serie), Csv(f.Numero), Csv(f.RazonSocial),
                    f.BaseImOpGra.ToString("0.00"), f.Exonerada.ToString("0.00"),
                    f.Inafecta.ToString("0.00"), f.IGV.ToString("0.00"),
                    f.OtrosTrib.ToString("0.00"), f.ImporteT.ToString("0.00"),
                    f.TipoCambio.ToString("0.00"), f.Descuento.ToString("0.00"), f.Recargo.ToString("0.00"))));
                break;
            case TipoReporteRegistroVenta.AgrupadoPorTipoDocumento:
                lineas = ["Voucher,FEmision,TDoc,Serie,Numero1,Numero2,RazonSocial,BaseImOpGra,IGV,ImporteT"];
                lineas.AddRange(resultado.FilasSunatAd.Select(f => string.Join(",",
                    Csv(f.Voucher), Csv(f.FEmision), Csv(f.TDoc),
                    Csv(f.Serie), Csv(f.Numero1), Csv(f.Numero2),
                    Csv(f.RazonSocial), f.BaseImOpGra.ToString("0.00"),
                    f.IGV.ToString("0.00"), f.ImporteT.ToString("0.00"))));
                break;
            case TipoReporteRegistroVenta.DetalladoPorComprobante:
                lineas = ["NumeroInterno,TipoDocumento,NumeroDocumento,Cliente,Descripcion,Cantidad,PrecioUnitario,IGV,Importe,Estado,FRegistro"];
                lineas.AddRange(resultado.FilasDetallado.Select(f => string.Join(",",
                    Csv(f.NumeroInterno), Csv(f.TipoDocumento), Csv(f.NumeroDocumento),
                    Csv(f.Cliente), Csv(f.Descripcion), f.Cantidad.ToString("0.##"),
                    f.PrecioUnitario.ToString("0.00"), f.IGV.ToString("0.00"),
                    f.Importe.ToString("0.00"), Csv(f.Estado), Csv(f.FRegistro))));
                break;
            default:
                lineas = ["FRegistro,TDocumento,TipoDocumento,TipoPago,EstadoDocumento,NNeto,NImpuesto1,NVenta,NRecargo,NDescuento"];
                lineas.AddRange(resultado.Filas.Select(f => string.Join(",",
                    f.FRegistro.ToString("dd/MM/yyyy HH:mm"), Csv(f.TDocumento),
                    Csv(f.TTipoDocumento), Csv(f.TipoPago), Csv(f.TEstadoDocumento),
                    f.NNeto.ToString("0.00"), f.NImpuesto1.ToString("0.00"),
                    f.NVenta.ToString("0.00"), f.NRecargo.ToString("0.00"), f.NDescuento.ToString("0.00"))));
                break;
        }

        File.WriteAllLines(dialog.FileName, lineas, System.Text.Encoding.UTF8);
        MessageBox.Show($"Exportado: {dialog.FileName}", "Exportar CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static string Csv(string valor)
        => $"\"{(valor ?? string.Empty).Replace("\"", "\"\"")}\"";

    private void SetBotones(bool habilitados)
    {
        foreach (Button btn in Controls.OfType<Button>())
            btn.Enabled = habilitados;
    }

    // ── InitializeComponent ────────────────────────────────────────────────────

    private void InitializeComponent()
    {
        Text = "Registro de Ventas";
        Size = new Size(850, 680);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9f);

        // ── Grupo Tipo de Reporte ──────────────────────────────────────────────
        var grpTipo = new GroupBox
        {
            Text = "Tipo de Reporte",
            Location = new Point(12, 8),
            Size = new Size(400, 240),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            BackColor = Color.White
        };

        _rdoSunat = MakeRadio("Correlativo de Documento Sunat", new Point(12, 24), true);
        _rdoEstado = MakeRadio("Estado de Documentos", new Point(12, 54));
        _rdoAgrupadoFechas = MakeRadio("Agrupado por Fechas", new Point(12, 84));
        _rdoAgrupadoTipoDoc = MakeRadio("Agrupado por Tipo Documento", new Point(12, 114));
        _rdoCorrelativo = MakeRadio("Correlativo de Documento", new Point(12, 144));
        _rdoDetalladoComprobante = MakeRadio("Detallado por Comprobante", new Point(12, 174));
        _rdoCorrelativoDetallado = MakeRadio("Correlativo de Documento Detallado", new Point(12, 204));

        grpTipo.Controls.AddRange(new Control[]
        {
            _rdoSunat, _rdoEstado, _rdoAgrupadoFechas, _rdoAgrupadoTipoDoc,
            _rdoCorrelativo, _rdoDetalladoComprobante, _rdoCorrelativoDetallado
        });

        // Nota: tipo 7 se deja como gap, habilitado como aviso
        _rdoConFormaPago = MakeRadio("Correlativo con Forma de Pago (GAP)", new Point(12, 234));
        grpTipo.Height = 265;
        grpTipo.Controls.Add(_rdoConFormaPago);

        foreach (RadioButton rdo in grpTipo.Controls.OfType<RadioButton>())
            rdo.CheckedChanged += RdoTipo_CheckedChanged;

        // ── Grupo Opciones ─────────────────────────────────────────────────────
        var grpOpciones = new GroupBox
        {
            Text = "Opciones",
            Location = new Point(424, 8),
            Size = new Size(404, 265),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            BackColor = Color.White
        };

        var lblOrden = MakeLabel("Orden:", new Point(12, 26));
        _cboOrden = new ComboBox
        {
            Location = new Point(130, 22),
            Width = 160,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 9f)
        };
        _cboOrden.Items.AddRange(["Correlativo", "Montos", "Fechas"]);
        _cboOrden.SelectedIndex = 0;

        var lblRedondeo = MakeLabel("Redondeo:", new Point(12, 56));
        _cboRedondeo = new ComboBox
        {
            Location = new Point(130, 52),
            Width = 160,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 9f)
        };
        _cboRedondeo.Items.AddRange(["Entero", "Dos Decimales"]);
        _cboRedondeo.SelectedIndex = 0;

        _chkDiaContable = new CheckBox
        {
            Text = "Evaluar por Día Contable",
            Location = new Point(12, 88),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };
        _chkDiaContable.CheckedChanged += ChkDiaContable_CheckedChanged;

        _chkRegistroVenta = new CheckBox
        {
            Text = "Solo documentos en Registro de Ventas",
            Location = new Point(12, 116),
            AutoSize = true,
            Checked = true,
            Font = new Font("Segoe UI", 9f)
        };

        _chkTransferenciaGratuita = new CheckBox
        {
            Text = "Visualizar valores por Transferencia Gratuita",
            Location = new Point(12, 144),
            AutoSize = true,
            Visible = true,
            Font = new Font("Segoe UI", 9f)
        };

        grpOpciones.Controls.AddRange(new Control[]
        {
            lblOrden, _cboOrden,
            lblRedondeo, _cboRedondeo,
            _chkDiaContable, _chkRegistroVenta, _chkTransferenciaGratuita
        });

        // ── Grupo Filtros ──────────────────────────────────────────────────────
        var grpFiltros = new GroupBox
        {
            Text = "Filtros",
            Location = new Point(12, 285),
            Size = new Size(816, 210),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            BackColor = Color.White
        };

        // Cliente
        var lblCliente = MakeLabel("Cliente:", new Point(12, 28));
        _txtCliente = new TextBox { Location = new Point(130, 24), Width = 160, Font = new Font("Segoe UI", 9f), ReadOnly = true };
        _btnBuscarCliente = new Button
        {
            Text = "...",
            Location = new Point(295, 23),
            Size = new Size(30, 23),
            Font = new Font("Segoe UI", 9f)
        };
        _btnBuscarCliente.Click += BtnBuscarCliente_Click;

        // TipoDocumento
        var lblTipoDoc = MakeLabel("Tipo Documento:", new Point(12, 62));
        _chkTodosDoc = new CheckBox { Text = "Todos", Location = new Point(450, 58), AutoSize = true, Checked = true, Font = new Font("Segoe UI", 9f) };
        _cboTipoDocumento = new ComboBox { Location = new Point(130, 58), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false, Font = new Font("Segoe UI", 9f) };
        _chkTodosDoc.CheckedChanged += ChkTodos_CheckedChanged;

        // Estado
        var lblEstadoDoc = MakeLabel("Estado:", new Point(12, 96));
        _chkTodosEstado = new CheckBox { Text = "Todos", Location = new Point(450, 92), AutoSize = true, Checked = true, Font = new Font("Segoe UI", 9f) };
        _cboEstado = new ComboBox { Location = new Point(130, 92), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false, Font = new Font("Segoe UI", 9f) };
        _chkTodosEstado.CheckedChanged += ChkTodos_CheckedChanged;

        // Caja
        var lblCaja = MakeLabel("Caja:", new Point(12, 130));
        _chkTodasCajas = new CheckBox { Text = "Todas", Location = new Point(450, 126), AutoSize = true, Checked = true, Font = new Font("Segoe UI", 9f) };
        _cboCaja = new ComboBox { Location = new Point(130, 126), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false, Font = new Font("Segoe UI", 9f) };
        _chkTodasCajas.CheckedChanged += ChkTodos_CheckedChanged;

        // TipoPago
        var lblTipoPago = MakeLabel("Tipo de Pago:", new Point(12, 164));
        _chkTodosPagos = new CheckBox { Text = "Todos", Location = new Point(450, 160), AutoSize = true, Checked = true, Font = new Font("Segoe UI", 9f) };
        _cboTipoPago = new ComboBox { Location = new Point(130, 160), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false, Font = new Font("Segoe UI", 9f) };
        _chkTodosPagos.CheckedChanged += ChkTodos_CheckedChanged;

        grpFiltros.Controls.AddRange(new Control[]
        {
            lblCliente, _txtCliente, _btnBuscarCliente,
            lblTipoDoc, _cboTipoDocumento, _chkTodosDoc,
            lblEstadoDoc, _cboEstado, _chkTodosEstado,
            lblCaja, _cboCaja, _chkTodasCajas,
            lblTipoPago, _cboTipoPago, _chkTodosPagos
        });

        // ── Grupo Fechas ───────────────────────────────────────────────────────
        var grpFechas = new GroupBox
        {
            Text = "Rango de Fechas",
            Location = new Point(12, 505),
            Size = new Size(816, 80),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            BackColor = Color.White
        };

        _dtpFecIni = new DateTimePicker { Location = new Point(130, 24), Width = 100, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _dtpHorIni = new DateTimePicker { Location = new Point(240, 24), Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Value = DateTime.Today };
        _dtpFecFin = new DateTimePicker { Location = new Point(450, 24), Width = 100, Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _dtpHorFin = new DateTimePicker { Location = new Point(560, 24), Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Value = DateTime.Today };

        grpFechas.Controls.AddRange(new Control[]
        {
            MakeLabel("Fecha Inicio:", new Point(12, 28)), _dtpFecIni, _dtpHorIni,
            MakeLabel("Fecha Fin:", new Point(340, 28)), _dtpFecFin, _dtpHorFin
        });

        // ── Panel Año/Mes (solo para AgrupadoPorFechas) ────────────────────────
        _pnlAnioMes = new Panel
        {
            Location = new Point(12, 505),
            Size = new Size(816, 80),
            BackColor = Color.FromArgb(240, 248, 255),
            BorderStyle = BorderStyle.FixedSingle,
            Visible = false
        };
        _nudAno = new NumericUpDown { Location = new Point(160, 24), Width = 70, Minimum = 2000, Maximum = 2099, Value = DateTime.Today.Year, Font = new Font("Segoe UI", 9f) };
        _cboMes = new ComboBox { Location = new Point(320, 24), Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 9f) };
        _cboMes.Items.AddRange(NombresMeses);
        _cboMes.SelectedIndex = DateTime.Today.Month - 1;
        _nudHoraCorte = new NumericUpDown { Location = new Point(570, 24), Width = 55, Minimum = 0, Maximum = 23, Value = 0, Font = new Font("Segoe UI", 9f) };

        _pnlAnioMes.Controls.AddRange(new Control[]
        {
            MakeLabel("Año:", new Point(12, 28)), _nudAno,
            MakeLabel("Mes:", new Point(250, 28)), _cboMes,
            MakeLabel("Hora Corte:", new Point(460, 28)), _nudHoraCorte
        });

        // ── Botones ────────────────────────────────────────────────────────────
        var btnEmitir = MakeButton("Vista Previa", new Point(12, 598), Color.FromArgb(0, 120, 212));
        var btnImprimir = MakeButton("Imprimir", new Point(130, 598), Color.FromArgb(16, 124, 16));
        var btnExportar = MakeButton("Exportar CSV", new Point(248, 598), Color.FromArgb(136, 0, 21));
        var btnCerrar = MakeButton("Cerrar", new Point(718, 598), Color.FromArgb(96, 96, 96));

        btnEmitir.Click += BtnEmitir_Click;
        btnImprimir.Click += BtnImprimir_Click;
        btnExportar.Click += BtnExportar_Click;
        btnCerrar.Click += BtnCerrar_Click;

        // ── Estado ─────────────────────────────────────────────────────────────
        _lblEstado = new Label
        {
            Location = new Point(12, 640),
            Size = new Size(816, 20),
            ForeColor = Color.FromArgb(0, 120, 212),
            Font = new Font("Segoe UI", 8.5f)
        };

        Load += FrmRepRegistroVentaReporte_Load;

        Controls.AddRange(new Control[]
        {
            grpTipo, grpOpciones, grpFiltros, grpFechas, _pnlAnioMes,
            btnEmitir, btnImprimir, btnExportar, btnCerrar,
            _lblEstado
        });
    }

    private static RadioButton MakeRadio(string text, Point loc, bool isChecked = false)
        => new()
        {
            Text = text,
            Location = loc,
            AutoSize = true,
            Checked = isChecked,
            Font = new Font("Segoe UI", 9f)
        };

    private static Label MakeLabel(string text, Point loc)
        => new() { Text = text, Location = loc, AutoSize = true, Font = new Font("Segoe UI", 9f) };

    private static Button MakeButton(string text, Point loc, Color color)
    {
        var btn = new Button
        {
            Text = text,
            Location = loc,
            Size = new Size(110, 36),
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    /// <summary>Item para combos con código + descripción.</summary>
    private sealed class FiltroItem
    {
        public string Codigo { get; }
        private readonly string _descripcion;
        public FiltroItem(string codigo, string descripcion) { Codigo = codigo; _descripcion = descripcion; }
        public override string ToString() => _descripcion;
    }
}
