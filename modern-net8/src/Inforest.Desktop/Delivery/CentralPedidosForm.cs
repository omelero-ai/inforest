using Inforest.Application.CentralPedidos;
using Inforest.Domain.Repositories;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario de Central de Pedidos multi-local.
/// <para>
/// Legacy: <c>frmCentralPedidos.frm</c> — "Correlativo Pedidos Centro Produccion".
/// Consulta pedidos de la base <c>CENTRALDELIVERY</c> mediante
/// <c>sp_Inforest_PedidosCentralPedido</c>.
/// Habilitado solo si flag <c>lCD = 1</c> en <c>TCAJA</c>.
/// </para>
/// Reglas: BR-DEL-003, BR-DEL-008.
/// </summary>
public partial class CentralPedidosForm : Form
{
    private readonly ObtenerPedidosCentralHandler _obtenerHandler;

    public CentralPedidosForm(ObtenerPedidosCentralHandler obtenerHandler)
    {
        _obtenerHandler = obtenerHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Central de Pedidos";
        Size = new Size(1100, 650);
        WindowState = FormWindowState.Maximized;

        var toolbar = new ToolStrip();
        var btnConsultar = new ToolStripButton("Consultar");
        var btnEmitir = new ToolStripButton("Emitir");
        toolbar.Items.AddRange([btnConsultar, btnEmitir]);

        var lblFechaIni = new Label { Text = "Desde:", Left = 10, Top = 35, Width = 50 };
        var dtpFecIni = new DateTimePicker { Left = 65, Top = 32, Width = 120, Value = DateTime.Today };
        var lblFechaFin = new Label { Text = "Hasta:", Left = 195, Top = 35, Width = 50 };
        var dtpFecFin = new DateTimePicker { Left = 250, Top = 32, Width = 120, Value = DateTime.Today };

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

        btnConsultar.Click += async (_, _) =>
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                var result = await _obtenerHandler.HandleAsync(
                    new ObtenerPedidosCentralQuery(dtpFecIni.Value.Date, dtpFecFin.Value.Date.AddHours(23).AddMinutes(59)));
                if (result.EsExitoso)
                    grid.DataSource = result.Valor?.ToList();
                else
                    MessageBox.Show(result.MensajeError, "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Cursor = Cursors.Default; }
        };

        Load += async (_, _) => btnConsultar.PerformClick();
    }
}
