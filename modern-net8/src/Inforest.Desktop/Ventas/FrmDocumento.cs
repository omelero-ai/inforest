using Inforest.Application.Caja;
using Inforest.Application.Configuracion;
using Inforest.Application.Maestros;
using Inforest.Application.Pedidos;
using Inforest.Application.Ventas;
using Inforest.Desktop.Caja;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Pantalla de gestión de documentos de venta (generación, cobro, reimpresión, anulación).
/// Legacy: frmDocumento.frm — "Generación de Documentos".
/// BR-DOC-001..010. POS-FUNC-006.
/// </summary>
public class FrmDocumento : Form
{
    // ── handlers ──────────────────────────────────────────────────────────────
    private readonly ObtenerItemsPendientesFacturacionHandler? _itemsHandler;
    private readonly ObtenerDocumentosPendientesCajaHandler? _documentosHandler;
    private readonly ObtenerPedidoPorCodigoHandler? _pedidoHandler;
    private readonly EmitirDocumentoHandler? _emitirHandler;
    private readonly AnularDocumentoHandler? _anularHandler;
    private readonly ReimprimirDocumentoHandler? _reimprimirHandler;
    private readonly ObtenerNotasCreditoPorFechaHandler? _ncObtenerHandler;
    private readonly AnularNotaCreditoHandler? _ncAnularHandler;
    private readonly EmitirNotaCreditoHandler? _ncEmitirHandler;
    private readonly ObtenerMediosPagoHandler? _mediosPagoHandler;
    private readonly PagarDocumentoHandler? _pagarHandler;
    private readonly RegistrarPagosMultiplesHandler? _registrarPagosMultiplesHandler;
    private readonly ProcesarPagoPinPadHandler? _procesarPagoPinPadHandler;
    private readonly ObtenerTerminalesPinPadHandler? _obtenerTerminalesPinPadHandler;
    private readonly ObtenerConfiguracionSistemaHandler? _configuracionHandler;
    private readonly ObtenerClientesActivosHandler? _clientesHandler;
    private readonly ObtenerClienteGeneralBoletaHandler? _clienteGeneralBoletaHandler;

    // ── parámetros de apertura ─────────────────────────────────────────────────
    private readonly string? _codigoPedido;
    private readonly string? _codigoCaja;
    private readonly string? _usuario;
    private readonly bool _permiteNotasCredito;

    // ── estado en memoria ──────────────────────────────────────────────────────
    private readonly List<ItemPendienteFacturacionDto> _itemsPendientes = [];
    private readonly List<ItemPendienteFacturacionDto> _itemsSeleccionados = [];
    private int _divisor = 1;

    // ── grids ──────────────────────────────────────────────────────────────────
    private readonly DataGridView _grdPendientes;
    private readonly DataGridView _grdSeleccionados;
    private readonly DataGridView _grdDocumentos;

    // ── totales ────────────────────────────────────────────────────────────────
    private readonly Label _lblTotalSel;
    private readonly Label _lblDivisor;
    private readonly Label _lblDivTotal;

    // ── botones de acción ──────────────────────────────────────────────────────
    private readonly Button _btnGenerar;
    private readonly Button _btnPagar;
    private readonly Button _btnReimprimir;
    private readonly Button _btnAnular;
    private readonly Button _btnNotasCredito;
    private readonly Button _btnSalir;

    public FrmDocumento(
        string? codigoPedido = null,
        string? codigoCaja = null,
        string? usuario = null,
        bool permiteNotasCredito = false,
        ObtenerItemsPendientesFacturacionHandler? itemsHandler = null,
        ObtenerDocumentosPendientesCajaHandler? documentosHandler = null,
        ObtenerPedidoPorCodigoHandler? pedidoHandler = null,
        EmitirDocumentoHandler? emitirHandler = null,
        AnularDocumentoHandler? anularHandler = null,
        ReimprimirDocumentoHandler? reimprimirHandler = null,
        ObtenerNotasCreditoPorFechaHandler? ncObtenerHandler = null,
        AnularNotaCreditoHandler? ncAnularHandler = null,
        EmitirNotaCreditoHandler? ncEmitirHandler = null,
        ObtenerMediosPagoHandler? mediosPagoHandler = null,
        PagarDocumentoHandler? pagarHandler = null,
        RegistrarPagosMultiplesHandler? registrarPagosMultiplesHandler = null,
        ProcesarPagoPinPadHandler? procesarPagoPinPadHandler = null,
        ObtenerTerminalesPinPadHandler? obtenerTerminalesPinPadHandler = null,
        ObtenerConfiguracionSistemaHandler? configuracionHandler = null,
        ObtenerClientesActivosHandler? clientesHandler = null,
        ObtenerClienteGeneralBoletaHandler? clienteGeneralBoletaHandler = null)
    {
        _codigoPedido = codigoPedido;
        _codigoCaja = codigoCaja;
        _usuario = usuario;
        _permiteNotasCredito = permiteNotasCredito;
        _itemsHandler = itemsHandler;
        _documentosHandler = documentosHandler;
        _pedidoHandler = pedidoHandler;
        _emitirHandler = emitirHandler;
        _anularHandler = anularHandler;
        _reimprimirHandler = reimprimirHandler;
        _ncObtenerHandler = ncObtenerHandler;
        _ncAnularHandler = ncAnularHandler;
        _ncEmitirHandler = ncEmitirHandler;
        _mediosPagoHandler = mediosPagoHandler;
        _pagarHandler = pagarHandler;
        _registrarPagosMultiplesHandler = registrarPagosMultiplesHandler;
        _procesarPagoPinPadHandler = procesarPagoPinPadHandler;
        _obtenerTerminalesPinPadHandler = obtenerTerminalesPinPadHandler;
        _configuracionHandler = configuracionHandler;
        _clientesHandler = clientesHandler;
        _clienteGeneralBoletaHandler = clienteGeneralBoletaHandler;

        Text = "Generación de Documentos";
        WindowState = FormWindowState.Maximized;
        StartPosition = FormStartPosition.CenterScreen;
        MinimizeBox = false;
        Font = new Font("Segoe UI", 9f);

        // ── grids ──────────────────────────────────────────────────────────────
        _grdPendientes = BuildGrid("Por Facturar");
        _grdSeleccionados = BuildGrid("Seleccionados");
        _grdDocumentos = BuildGrid("Documentos por Cobrar");

        ConfigurarColumnasPendientes(_grdPendientes);
        ConfigurarColumnasPendientes(_grdSeleccionados);
        ConfigurarColumnasDocumentos(_grdDocumentos);

        _grdPendientes.CellDoubleClick += (_, _) => MoverUno(_grdPendientes, _itemsPendientes, _itemsSeleccionados);
        _grdSeleccionados.CellDoubleClick += (_, _) => MoverUno(_grdSeleccionados, _itemsSeleccionados, _itemsPendientes);
        _grdDocumentos.SelectionChanged += GrdDocumentos_SelectionChanged;

        // ── botones de movimiento ──────────────────────────────────────────────
        var btnPasarTodos = new Button { Text = "▶▶ Todos", Dock = DockStyle.Top, Height = 30 };
        var btnPasarUno = new Button { Text = "▶ Uno", Dock = DockStyle.Top, Height = 30 };
        var btnQuitarUno = new Button { Text = "◀ Uno", Dock = DockStyle.Top, Height = 30 };
        var btnQuitarTodos = new Button { Text = "◀◀ Todos", Dock = DockStyle.Top, Height = 30 };

        btnPasarTodos.Click += (_, _) => { _itemsSeleccionados.AddRange(_itemsPendientes); _itemsPendientes.Clear(); RefrescarGridsItems(); };
        btnPasarUno.Click += (_, _) => MoverUno(_grdPendientes, _itemsPendientes, _itemsSeleccionados);
        btnQuitarUno.Click += (_, _) => MoverUno(_grdSeleccionados, _itemsSeleccionados, _itemsPendientes);
        btnQuitarTodos.Click += (_, _) => { _itemsPendientes.AddRange(_itemsSeleccionados); _itemsSeleccionados.Clear(); RefrescarGridsItems(); };

        var panelMov = new Panel { Width = 90, Padding = new Padding(4) };
        panelMov.Controls.AddRange([btnQuitarTodos, btnQuitarUno, btnPasarUno, btnPasarTodos]);

        // ── panel divisor ──────────────────────────────────────────────────────
        _lblTotalSel = new Label { Text = "0.00", Font = new Font("Segoe UI", 11, FontStyle.Bold), TextAlign = ContentAlignment.MiddleRight, Width = 130 };
        _lblDivisor = new Label { Text = "1", Font = new Font("Segoe UI", 11, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter, Width = 40 };
        _lblDivTotal = new Label { Text = "0.00", Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.DarkRed, TextAlign = ContentAlignment.MiddleRight, Width = 130 };

        var btnMas = new Button { Text = "+", Width = 30, Height = 28 };
        var btnMenos = new Button { Text = "−", Width = 30, Height = 28 };
        btnMas.Click += (_, _) => { _divisor++; ActualizarDivisor(); };
        btnMenos.Click += (_, _) => { if (_divisor > 1) { _divisor--; ActualizarDivisor(); } };

        var panelDiv = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Padding = new Padding(4) };
        panelDiv.Controls.AddRange([
            new Label { Text = "Total:", TextAlign = ContentAlignment.MiddleLeft, Width = 40, Height = 28 },
            _lblTotalSel,
            new Label { Text = "÷", TextAlign = ContentAlignment.MiddleCenter, Width = 20, Height = 28 },
            btnMenos, _lblDivisor, btnMas,
            new Label { Text = "=", TextAlign = ContentAlignment.MiddleCenter, Width = 20, Height = 28 },
            _lblDivTotal
        ]);

        // ── botones de acción ──────────────────────────────────────────────────
        _btnGenerar = ActionButton("Generar Documento", Color.FromArgb(0, 120, 215));
        _btnPagar = ActionButton("Pagos del Documento", Color.FromArgb(16, 124, 16));
        _btnReimprimir = ActionButton("Re-Imprimir", Color.FromArgb(93, 64, 55));
        _btnAnular = ActionButton("Anulación Documento", Color.FromArgb(196, 43, 28));
        _btnNotasCredito = ActionButton("Notas de Crédito", Color.FromArgb(70, 70, 130));
        _btnSalir = ActionButton("Salir", Color.FromArgb(68, 68, 68));

        _btnNotasCredito.Visible = _permiteNotasCredito;

        _btnGenerar.Click += async (_, _) => await GenerarDocumentoAsync();
        _btnPagar.Click += async (_, _) => await PagarDocumentoAsync();
        _btnReimprimir.Click += async (_, _) => await ReimprimirDocumentoAsync();
        _btnAnular.Click += async (_, _) => await AnularDocumentoAsync();
        _btnNotasCredito.Click += (_, _) => AbrirNotasCredito();
        _btnSalir.Click += (_, _) => Close();

        var panelBotones = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            Padding = new Padding(4),
            FlowDirection = FlowDirection.LeftToRight
        };
        panelBotones.Controls.AddRange([
            _btnGenerar, _btnPagar, _btnReimprimir, _btnAnular, _btnNotasCredito, _btnSalir
        ]);

        // ── layout superior (items + movimiento) ───────────────────────────────
        var splitterSup = new SplitContainer
        {
            Orientation = Orientation.Vertical,
            SplitterDistance = 420,
            Dock = DockStyle.Fill
        };
        splitterSup.Panel1.Controls.Add(_grdPendientes);
        splitterSup.Panel2.Controls.Add(_grdSeleccionados);

        var panelSup = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
        panelSup.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 46));
        panelSup.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
        panelSup.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 54));
        panelSup.Controls.Add(_grdPendientes, 0, 0);
        panelSup.Controls.Add(panelMov, 1, 0);
        panelSup.Controls.Add(_grdSeleccionados, 2, 0);

        // ── layout principal ───────────────────────────────────────────────────
        var splitterMain = new SplitContainer
        {
            Orientation = Orientation.Horizontal,
            Dock = DockStyle.Fill
        };
        splitterMain.Panel1.Controls.Add(panelSup);
        splitterMain.Panel2.Controls.Add(_grdDocumentos);
        splitterMain.SplitterDistance = 300;

        var panelDivRow = new Panel { Dock = DockStyle.Bottom, Height = 40 };
        panelDivRow.Controls.Add(panelDiv);
        splitterMain.Panel1.Controls.Add(panelDivRow);

        Controls.Add(splitterMain);
        Controls.Add(panelBotones);

        Load += async (_, _) => await CargarDatosAsync();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Carga inicial de datos
    // ─────────────────────────────────────────────────────────────────────────

    private async Task CargarDatosAsync()
    {
        if (_itemsHandler is not null && !string.IsNullOrWhiteSpace(_codigoPedido))
        {
            var items = await _itemsHandler.HandleAsync(
                new ObtenerItemsPendientesFacturacionQuery(_codigoPedido));
            _itemsPendientes.Clear();
            _itemsPendientes.AddRange(items);
            RefrescarGridsItems();
        }

        if (_documentosHandler is not null && !string.IsNullOrWhiteSpace(_codigoCaja))
        {
            var docs = await _documentosHandler.HandleAsync(
                new ObtenerDocumentosPendientesCajaQuery(_codigoCaja));
            RefrescarGridDocumentos(docs);
        }

        ActualizarBotones();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Acciones de ítems (BR-DOC-002)
    // ─────────────────────────────────────────────────────────────────────────

    private void MoverUno(DataGridView origen, List<ItemPendienteFacturacionDto> listaOrigen, List<ItemPendienteFacturacionDto> listaDestino)
    {
        if (origen.CurrentRow?.DataBoundItem is not ItemPendienteFacturacionDto item) return;
        listaOrigen.Remove(item);
        listaDestino.Add(item);
        RefrescarGridsItems();
    }

    private void RefrescarGridsItems()
    {
        _grdPendientes.DataSource = null;
        _grdPendientes.DataSource = _itemsPendientes.Select(i => new
        {
            i.CodigoPedido,
            Producto = i.NombreProducto,
            Precio = i.PrecioVenta,
            Cant = i.Cantidad,
            SubTotal = i.Venta
        }).ToList();

        _grdSeleccionados.DataSource = null;
        _grdSeleccionados.DataSource = _itemsSeleccionados.Select(i => new
        {
            i.CodigoPedido,
            Producto = i.NombreProducto,
            Precio = i.PrecioVenta,
            Cant = i.Cantidad,
            SubTotal = i.Venta
        }).ToList();

        ActualizarDivisor();
        ActualizarBotones();
    }

    private void ActualizarDivisor()
    {
        var total = _itemsSeleccionados.Sum(i => i.Venta);
        _lblTotalSel.Text = total.ToString("N2");
        _lblDivisor.Text = _divisor.ToString();
        _lblDivTotal.Text = (_divisor > 0 ? total / _divisor : 0).ToString("N2");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Refresh lista documentos
    // ─────────────────────────────────────────────────────────────────────────

    private void RefrescarGridDocumentos(IReadOnlyList<DocumentoPendienteDto> docs)
    {
        _grdDocumentos.DataSource = null;
        _grdDocumentos.DataSource = docs.Select(d => new
        {
            Documento = d.CodigoDocumento,
            Mesa = d.Mesa ?? "",
            Observacion = d.Observacion ?? "",
            Pedido = d.CodigoPedido ?? "",
            d.Total,
            Fecha = d.FechaRegistro.ToString("dd-MMM"),
            Cliente = d.Cliente ?? ""
        }).ToList();
    }

    private DocumentoPendienteDto? DocumentoSeleccionado =>
        _grdDocumentos.CurrentRow?.Cells["Documento"]?.Value is string cod
            ? null // populated from async reload below
            : null;

    private string? CodigoDocumentoSeleccionado =>
        _grdDocumentos.CurrentRow?.Cells["Documento"]?.Value as string;

    private decimal TotalDocumentoSeleccionado =>
        _grdDocumentos.CurrentRow?.Cells["Total"]?.Value is decimal d ? d : 0m;

    private void GrdDocumentos_SelectionChanged(object? sender, EventArgs e) => ActualizarBotones();

    // ─────────────────────────────────────────────────────────────────────────
    // Generar Documento — BR-DOC-003 (abre FrmVenta con pedido cargado)
    // Legacy: frmDocumento CmdOpcion 2 — frmGeneracion.Show vbModal
    // ─────────────────────────────────────────────────────────────────────────

    private async Task GenerarDocumentoAsync()
    {
        if (_itemsSeleccionados.Count == 0)
        {
            MessageBox.Show("No existen ítems a facturar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_pedidoHandler is null || _emitirHandler is null)
        {
            MessageBox.Show("Función de generación de documento no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var codigoPedido = _itemsSeleccionados.FirstOrDefault()?.CodigoPedido ?? _codigoPedido ?? "";
        var result = await _pedidoHandler.HandleAsync(new ObtenerPedidoPorCodigoQuery(codigoPedido));
        if (!result.EsExitoso || result.Valor is null)
        {
            MessageBox.Show("No se pudo cargar el pedido.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var frmVenta = new FrmVenta(
            result.Valor,
            _emitirHandler,
            _mediosPagoHandler,
            _pagarHandler,
            _registrarPagosMultiplesHandler,
            _procesarPagoPinPadHandler,
            _obtenerTerminalesPinPadHandler,
            _anularHandler,
            _configuracionHandler,
            _clientesHandler,
            _clienteGeneralBoletaHandler);

        if (frmVenta.ShowDialog(this) == DialogResult.OK)
            await CargarDatosAsync();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Pagar Documento — BR-DOC-004 (frmPago)
    // Legacy: frmDocumento CmdOpcion 6 — frmPago.Show vbModal
    // ─────────────────────────────────────────────────────────────────────────

    private async Task PagarDocumentoAsync()
    {
        var codDoc = CodigoDocumentoSeleccionado;
        if (string.IsNullOrWhiteSpace(codDoc)) return;

        var total = TotalDocumentoSeleccionado;

        using var frmPago = new FrmPago(
            codDoc,
            total,
            _codigoCaja,
            _mediosPagoHandler,
            _pagarHandler,
            _registrarPagosMultiplesHandler);

        if (frmPago.ShowDialog(this) == DialogResult.OK)
            await CargarDatosAsync();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Reimprimir Documento — BR-DOC-005
    // Legacy: frmDocumento CmdOpcion 7 — EXEC usp_Inforest_Impresion @doc, 3
    // ─────────────────────────────────────────────────────────────────────────

    private async Task ReimprimirDocumentoAsync()
    {
        var codDoc = CodigoDocumentoSeleccionado;
        if (string.IsNullOrWhiteSpace(codDoc)) return;

        if (_reimprimirHandler is null)
        {
            MessageBox.Show("Función de reimpresión no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var resp = MessageBox.Show(
            $"¿Desea reimprimir el documento {codDoc}?",
            Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (resp != DialogResult.Yes) return;

        var result = await _reimprimirHandler.HandleAsync(new ReimprimirDocumentoCommand(codDoc, 3));
        if (!result.EsExitoso)
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        else
            MessageBox.Show("Documento reenviado a impresora.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Anular Documento — BR-DOC-006
    // Legacy: frmDocumento CmdOpcion 9 — AnularDocumentoHandler
    // ─────────────────────────────────────────────────────────────────────────

    private async Task AnularDocumentoAsync()
    {
        var codDoc = CodigoDocumentoSeleccionado;
        if (string.IsNullOrWhiteSpace(codDoc)) return;

        if (_anularHandler is null)
        {
            MessageBox.Show("Función de anulación no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var dlg = new Form
        {
            Text = "Motivo de Anulación",
            Width = 400,
            Height = 160,
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false
        };
        var txtMotivo = new TextBox { Left = 16, Top = 20, Width = 352 };
        var btnOk = new Button { Left = 130, Top = 68, Width = 100, Text = "Anular", DialogResult = DialogResult.OK };
        var btnCancel = new Button { Left = 240, Top = 68, Width = 100, Text = "Cancelar", DialogResult = DialogResult.Cancel };
        dlg.Controls.AddRange([txtMotivo, btnOk, btnCancel]);
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCancel;

        if (dlg.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(txtMotivo.Text))
            return;

        var result = await _anularHandler.HandleAsync(
            new AnularDocumentoCommand(codDoc, _usuario ?? Environment.UserName, txtMotivo.Text.Trim()));

        if (!result.EsExitoso)
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        else
        {
            MessageBox.Show("Documento anulado correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            await CargarDatosAsync();
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Notas de Crédito — BR-DOC-007 (condicional TPARAMETRO.lAnula)
    // Legacy: frmDocumento cmdNotasCredito 13 — frmNotaCreditoDetalle.Show vbModal
    // ─────────────────────────────────────────────────────────────────────────

    private void AbrirNotasCredito()
    {
        if (_ncObtenerHandler is null || _ncAnularHandler is null || _ncEmitirHandler is null)
        {
            MessageBox.Show("Función de notas de crédito no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var frmNC = new FrmNotaCredito(
            _ncObtenerHandler,
            _ncAnularHandler,
            _ncEmitirHandler,
            _usuario ?? Environment.UserName);
        frmNC.ShowDialog(this);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Estado de botones
    // ─────────────────────────────────────────────────────────────────────────

    private void ActualizarBotones()
    {
        var hayItemsSel = _itemsSeleccionados.Count > 0;
        var hayDocSel = !string.IsNullOrWhiteSpace(CodigoDocumentoSeleccionado);

        _btnGenerar.Enabled = hayItemsSel;
        _btnPagar.Enabled = hayDocSel;
        _btnReimprimir.Enabled = hayDocSel;
        _btnAnular.Enabled = hayDocSel;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Helpers de UI
    // ─────────────────────────────────────────────────────────────────────────

    private static Button ActionButton(string text, Color backColor)
    {
        return new Button
        {
            Text = text,
            Width = 148,
            Height = 38,
            BackColor = backColor,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            Margin = new Padding(4)
        };
    }

    private static DataGridView BuildGrid(string caption)
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = Color.FromArgb(37, 37, 37),
                ForeColor = Color.White
            },
            DefaultCellStyle = new DataGridViewCellStyle { Font = new Font("Segoe UI", 8.5f) },
            AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(240, 245, 255)
            },
            RowHeadersVisible = false,
            BorderStyle = BorderStyle.None,
            GridColor = Color.LightSteelBlue
        };

        return grid;
    }

    private static void ConfigurarColumnasPendientes(DataGridView grid)
    {
        grid.AutoGenerateColumns = true;
    }

    private static void ConfigurarColumnasDocumentos(DataGridView grid)
    {
        grid.AutoGenerateColumns = true;
    }
}
