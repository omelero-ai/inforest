using Inforest.Application.Despacho;
using Inforest.Application.Delivery;
using Inforest.Domain.Entities.Delivery;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario del Despachador — panel de pedidos delivery en despacho.
/// <para>
/// Legacy: <c>mdiDespachador.frm</c> + <c>frmDespachador.frm</c>.
/// Muestra grilla de pedidos de la vista <c>vDespachador</c> con filtros de fecha,
/// ordenamiento por columna y acciones emitir/exportar.
/// </para>
/// Reglas: BR-DEL-003, BR-DEL-004.
/// </summary>
public partial class DespachadorForm : Form
{
    private readonly ObtenerPedidosDespachadorHandler _obtenerHandler;
    private readonly EmitirDespachoPedidoHandler _emitirHandler;
    private readonly ExportarDespachadorHandler _exportarHandler;

    private IEnumerable<PedidoDelivery> _pedidosActuales = [];

    public DespachadorForm(
        ObtenerPedidosDespachadorHandler obtenerHandler,
        EmitirDespachoPedidoHandler emitirHandler,
        ExportarDespachadorHandler exportarHandler)
    {
        _obtenerHandler = obtenerHandler;
        _emitirHandler = emitirHandler;
        _exportarHandler = exportarHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Despachador";
        Size = new Size(1200, 700);
        WindowState = FormWindowState.Maximized;

        var toolbar = new ToolStrip();
        var btnRefrescar = new ToolStripButton("Refrescar");
        var btnEmitir = new ToolStripButton("Emitir");
        var btnExportar = new ToolStripButton("Exportar");
        toolbar.Items.AddRange([btnRefrescar, btnEmitir, btnExportar]);

        var lblFechaIni = new Label { Text = "Desde:", Left = 10, Top = 35, Width = 50 };
        var dtpFecIni = new DateTimePicker { Left = 65, Top = 32, Width = 120, Name = "dtpFecIni", Value = DateTime.Today };
        var lblFechaFin = new Label { Text = "Hasta:", Left = 195, Top = 35, Width = 50 };
        var dtpFecFin = new DateTimePicker { Left = 250, Top = 32, Width = 120, Name = "dtpFecFin", Value = DateTime.Today };

        var grid = new DataGridView
        {
            Name = "grdPedidos",
            Left = 0, Top = 65,
            Width = ClientSize.Width,
            Height = ClientSize.Height - 65,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        Controls.AddRange([toolbar, lblFechaIni, dtpFecIni, lblFechaFin, dtpFecFin, grid]);

        btnRefrescar.Click += async (_, _) => await CargarPedidosAsync(dtpFecIni, dtpFecFin, grid);
        btnEmitir.Click += async (_, _) => await EmitirSeleccionadoAsync(grid);
        btnExportar.Click += async (_, _) => await ExportarAsync(dtpFecIni, dtpFecFin);
        Load += async (_, _) => await CargarPedidosAsync(dtpFecIni, dtpFecFin, grid);
    }

    private async Task CargarPedidosAsync(DateTimePicker ini, DateTimePicker fin, DataGridView grid)
    {
        Cursor = Cursors.WaitCursor;
        try
        {
            var result = await _obtenerHandler.HandleAsync(
                new ObtenerPedidosDespachadorQuery(ini.Value.Date, fin.Value.Date.AddHours(23).AddMinutes(59)));
            if (result.EsExitoso)
            {
                _pedidosActuales = result.Valor ?? [];
                grid.DataSource = _pedidosActuales.ToList();
            }
        }
        finally { Cursor = Cursors.Default; }
    }

    private async Task EmitirSeleccionadoAsync(DataGridView grid)
    {
        if (grid.SelectedRows.Count == 0) return;
        var pedido = (PedidoDelivery)grid.SelectedRows[0].DataBoundItem;
        var result = await _emitirHandler.HandleAsync(new EmitirDespachoPedidoCommand(pedido.CodigoPedido, "SISTEMA"));
        if (!result.EsExitoso)
            MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    private async Task ExportarAsync(DateTimePicker ini, DateTimePicker fin)
    {
        using var dlg = new SaveFileDialog { Filter = "HTML|*.html" };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        var result = await _exportarHandler.HandleAsync(
            new ExportarDespachadorQuery(ini.Value.Date, fin.Value.Date.AddHours(23).AddMinutes(59)));
        if (result.EsExitoso)
        {
            // Genera HTML simple con los datos exportados
            var html = GenerarHtml(result.Valor ?? []);
            await File.WriteAllTextAsync(dlg.FileName, html);
            MessageBox.Show("Exportación completada.", "Exportar", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private static string GenerarHtml(IEnumerable<PedidoDelivery> pedidos)
    {
        var rows = string.Join("\n", pedidos.Select(p =>
            $"<tr><td>{p.CodigoPedido}</td><td>{p.CodigoClienteDelivery}</td>" +
            $"<td>{p.FechaRegistro:dd/MM/yyyy HH:mm}</td><td>{p.EstadoDelivery}</td></tr>"));
        return $"<html><body><table border='1'><tr><th>Pedido</th><th>Cliente</th><th>Fecha</th><th>Estado</th></tr>{rows}</table></body></html>";
    }
}
