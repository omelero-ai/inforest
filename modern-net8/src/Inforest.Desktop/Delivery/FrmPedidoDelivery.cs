using Inforest.Application.Delivery;
using Inforest.Domain.Entities.Delivery;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario de Seguimiento de Pedidos Delivery.
/// <para>
/// Legacy: <c>frmPedidoDelivery.frm</c> — "Seguimientos de Pedidos Delivery".
/// Muestra pedidos pendientes de delivery desde <c>vDespachador</c>
/// donde <c>tTipoPedido='02'</c>, <c>tEstadoPedido='02'</c> y <c>fLlegada IS NULL</c>.
/// Permite filtrar por columna (teléfono, cliente, motorizado, etc.).
/// Visualiza el detalle de ítems del pedido seleccionado.
/// </para>
/// Legacy: <c>frmPedidoDelivery.frm</c>, <c>vDespachador</c>, <c>vPedidoDetalle</c>.
/// Reglas: BR-DEL-009.
/// </summary>
public partial class FrmPedidoDelivery : Form
{
    private readonly ObtenerPedidosSeguimientoDeliveryHandler _seguimientoHandler;
    private IReadOnlyList<PedidoDelivery> _pedidos = [];

    public FrmPedidoDelivery(ObtenerPedidosSeguimientoDeliveryHandler seguimientoHandler)
    {
        _seguimientoHandler = seguimientoHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = " Seguimientos de Pedidos Delivery ";
        Size = new Size(1200, 700);
        WindowState = FormWindowState.Maximized;

        // ── Toolbar ──────────────────────────────────────────────────────────
        var toolbar = new ToolStrip { Dock = DockStyle.Top };
        var btnActualizar = new ToolStripButton("Actualizar") { ToolTipText = "Recargar pedidos pendientes" };
        var btnSalir = new ToolStripButton("Salir");
        toolbar.Items.AddRange([btnActualizar, new ToolStripSeparator(), btnSalir]);

        // ── Filtro ───────────────────────────────────────────────────────────
        var panelFiltro = new Panel { Dock = DockStyle.Top, Height = 35 };
        var lblFiltro = new Label { Text = "Buscar:", Left = 5, Top = 8, Width = 55 };
        var cboCriterio = new ComboBox
        {
            Name = "cboCriterio",
            Left = 60, Top = 5, Width = 140,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        cboCriterio.Items.AddRange(["Teléfono", "Cliente", "Motorizado", "Pedido"]);
        cboCriterio.SelectedIndex = 0;
        var txtBusqueda = new TextBox { Name = "txtBusqueda", Left = 210, Top = 7, Width = 180 };
        var btnFiltrar = new Button { Text = "Filtrar", Left = 400, Top = 5, Width = 70 };
        panelFiltro.Controls.AddRange([lblFiltro, cboCriterio, txtBusqueda, btnFiltrar]);

        // ── Grid principal (cabecera) ─────────────────────────────────────────
        var gridCabecera = new DataGridView
        {
            Name = "grdCabecera",
            Dock = DockStyle.Top,
            Height = 260,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        };
        gridCabecera.Columns.AddRange(
            new DataGridViewTextBoxColumn { Name = "colPedido", HeaderText = "Pedido", DataPropertyName = "CodigoPedido", Width = 110 },
            new DataGridViewTextBoxColumn { Name = "colFecha", HeaderText = "Fecha", DataPropertyName = "FechaRegistro", Width = 120 },
            new DataGridViewTextBoxColumn { Name = "colCliente", HeaderText = "Cliente", DataPropertyName = "CodigoClienteDelivery", Width = 160 },
            new DataGridViewTextBoxColumn { Name = "colMotorizado", HeaderText = "Motorizado", DataPropertyName = "CodigoMotorizado", Width = 140 },
            new DataGridViewTextBoxColumn { Name = "colEstado", HeaderText = "Estado", DataPropertyName = "EstadoDelivery", Width = 100 },
            new DataGridViewTextBoxColumn { Name = "colMonto", HeaderText = "Monto", DataPropertyName = "MontoTotal", Width = 90 }
        );

        // ── Grid detalle ──────────────────────────────────────────────────────
        var lblDetalle = new Label
        {
            Text = "Detalle del Pedido",
            Dock = DockStyle.Top,
            Height = 20,
            Font = new Font(Font, FontStyle.Bold)
        };
        var gridDetalle = new DataGridView
        {
            Name = "grdDetalle",
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        };
        gridDetalle.Columns.AddRange(
            new DataGridViewTextBoxColumn { HeaderText = "Ítem", Width = 60 },
            new DataGridViewTextBoxColumn { HeaderText = "Producto", Width = 250 },
            new DataGridViewTextBoxColumn { HeaderText = "P.Venta", Width = 90 },
            new DataGridViewTextBoxColumn { HeaderText = "Cantidad", Width = 80 },
            new DataGridViewTextBoxColumn { HeaderText = "SubTotal", Width = 100 }
        );

        // ── Status bar ────────────────────────────────────────────────────────
        var statusBar = new StatusStrip();
        var lblStatus = new ToolStripStatusLabel { Name = "lblStatus", Text = "Listo" };
        statusBar.Items.Add(lblStatus);

        Controls.AddRange([statusBar, gridDetalle, lblDetalle, gridCabecera, panelFiltro, toolbar]);

        // ── Eventos ───────────────────────────────────────────────────────────
        Load += async (_, _) => await CargarPedidosAsync(gridCabecera, lblStatus);

        btnActualizar.Click += async (_, _) => await CargarPedidosAsync(gridCabecera, lblStatus);

        btnSalir.Click += (_, _) => Close();

        btnFiltrar.Click += (_, _) =>
        {
            var criterio = cboCriterio.Text;
            var texto = txtBusqueda.Text.Trim().ToLowerInvariant();
            if (string.IsNullOrEmpty(texto))
            {
                gridCabecera.DataSource = _pedidos.ToList();
                return;
            }

            var filtrados = criterio switch
            {
                "Teléfono" => _pedidos.Where(p => (p.CodigoClienteDelivery ?? "").Contains(texto, StringComparison.OrdinalIgnoreCase)),
                "Cliente" => _pedidos.Where(p => (p.CodigoClienteDelivery ?? "").Contains(texto, StringComparison.OrdinalIgnoreCase)),
                "Motorizado" => _pedidos.Where(p => (p.CodigoMotorizado ?? "").Contains(texto, StringComparison.OrdinalIgnoreCase)),
                "Pedido" => _pedidos.Where(p => p.CodigoPedido.Contains(texto, StringComparison.OrdinalIgnoreCase)),
                _ => _pedidos
            };
            gridCabecera.DataSource = filtrados.ToList();
            lblStatus.Text = $"Registro 0 de {filtrados.Count()}";
        };

        gridCabecera.SelectionChanged += (_, _) =>
        {
            // Legacy: Asignar() — muestra detalle del pedido seleccionado (vPedidoDetalle)
            // En .NET mostramos los campos disponibles en la entidad
            gridDetalle.Rows.Clear();
            if (gridCabecera.CurrentRow?.DataBoundItem is PedidoDelivery pedido)
            {
                lblStatus.Text = $"Pedido: {pedido.CodigoPedido} | Estado: {pedido.EstadoDelivery}";
            }
        };
    }

    private async Task CargarPedidosAsync(DataGridView grid, ToolStripStatusLabel status)
    {
        Cursor = Cursors.WaitCursor;
        try
        {
            var result = await _seguimientoHandler.HandleAsync(new ObtenerPedidosSeguimientoDeliveryQuery());
            if (result.EsExitoso && result.Valor is not null)
            {
                _pedidos = result.Valor.ToList();
                grid.DataSource = _pedidos.ToList();
                status.Text = $"Registro 0 de {_pedidos.Count}";
            }
            else
            {
                status.Text = result.MensajeError ?? "Error al cargar pedidos.";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }
}
